import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

// --------------------------------------------------
// Types
// --------------------------------------------------

export interface IPickedPerson {
  /** SharePoint user ID in this site (set when ensureUser succeeds). */
  id?: number;
  loginName: string;
  displayName: string;
  email?: string;
  text?: string;
  secondaryText?: string;
}

export interface ISharePointPeoplePickerProps {
  context: WebPartContext;
  value: IPickedPerson[];
  limit: number;
  /** 1 = user, 2 = distribution list, 4 = security group, 8 = SharePoint group */
  principalTypes?: number[];
  placeholder?: string;
  /** Resolve the user to a SharePoint user ID on selection. Default: true. */
  ensureUser?: boolean;
  /** Delay (ms) before searching while typing. Default: 300. */
  resolveDelay?: number;
  ariaLabel?: string;
  onChange: (people: IPickedPerson[]) => void;
  /** Show each person's profile picture. Default: false. */
  showProfilePicture?: boolean;
  /** Profile picture size in pixels. Default: 28. */
  profilePictureSize?: number;
}

interface IPickerEntity {
  Key?: string;
  DisplayText?: string;
  IsResolved?: boolean;
  Description?: string;
  EntityData?: {
    Email?: string;
    AccountName?: string;
  };
}

interface IPickerSearchResponse {
  value?: string;
  d?: {
    ClientPeoplePickerSearchUser?: string;
  };
}

// --------------------------------------------------
// Styles (inline so no SCSS changes are needed)
// --------------------------------------------------

const boxStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  alignItems: 'center',
  gap: 6,
  minHeight: 34,
  padding: '4px 8px',
  border: '1px solid #8a8886',
  borderRadius: 2,
  background: '#fff',
  cursor: 'text'
};

const chipStyle: React.CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 4,
  padding: '2px 4px 2px 8px',
  borderRadius: 12,
  background: '#e1dfdd',
  fontSize: 13
};

const chipButtonStyle: React.CSSProperties = {
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  fontSize: 16,
  lineHeight: 1,
  padding: '0 4px'
};

const inputStyle: React.CSSProperties = {
  flex: 1,
  minWidth: 140,
  border: 'none',
  outline: 'none',
  fontSize: 14,
  padding: '4px 0',
  background: 'transparent'
};

const listStyle: React.CSSProperties = {
  position: 'absolute',
  left: 0,
  right: 0,
  zIndex: 1000,
  margin: '2px 0 0',
  padding: 0,
  listStyle: 'none',
  maxHeight: 260,
  overflowY: 'auto',
  background: '#fff',
  border: '1px solid #c8c6c4',
  boxShadow: '0 3px 8px rgba(0,0,0,0.2)'
};

const itemStyle: React.CSSProperties = {
  padding: '8px 12px',
  cursor: 'pointer',
  fontSize: 14
};

const infoStyle: React.CSSProperties = {
  padding: '8px 12px',
  fontSize: 13,
  color: '#605e5c'
};

// --------------------------------------------------
// Component
// --------------------------------------------------

interface IProfilePictureProps {
  webUrl: string;
  displayName: string;
  email?: string;
  loginName?: string;
  size?: number;
}

/**
 * Gets the account name used by SharePoint's photo service.
 * Prefers the email; otherwise reads it from a claims login name
 * such as "i:0#.f|membership|user@company.com".
 */
const getAccountName = (email?: string, loginName?: string): string => {
  if (email) {
    return email;
  }

  if (loginName) {
    const parts = loginName.split('|');
    const last = parts[parts.length - 1];

    if (last && last.indexOf('@') > -1) {
      return last;
    }
  }

  return '';
};

/**
 * The person's SharePoint profile picture (no initials or placeholder).
 * Renders nothing if there is no account name or the picture fails to load.
 */
const ProfilePicture: React.FC<IProfilePictureProps> = (props) => {
  const size = props.size || 28;
  const [failed, setFailed] = React.useState<boolean>(false);

  const accountName = getAccountName(props.email, props.loginName);

  React.useEffect(() => {
    setFailed(false);
  }, [accountName]);

  if (!accountName || failed) {
    return null;
  }

  // SharePoint photo sizes: S (48px), M (72px), L (200px).
  const photoSize = size <= 48 ? 'S' : size <= 72 ? 'M' : 'L';

  const url =
    `${props.webUrl}/_layouts/15/userphoto.aspx` +
    `?size=${photoSize}&accountname=${encodeURIComponent(accountName)}`;

  return (
    <img
      src={url}
      alt={props.displayName}
      title={props.displayName}
      width={size}
      height={size}
      style={{
        width: size,
        height: size,
        borderRadius: '50%',
        objectFit: 'cover',
        flexShrink: 0
      }}
      onError={() => setFailed(true)}
    />
  );
};

const SharePointPeoplePicker: React.FC<ISharePointPeoplePickerProps> = (props) => {
  const { context, value, limit, onChange } = props;

  const [query, setQuery] = React.useState<string>('');
  const [suggestions, setSuggestions] = React.useState<IPickedPerson[]>([]);
  const [open, setOpen] = React.useState<boolean>(false);
  const [busy, setBusy] = React.useState<boolean>(false);
  const [message, setMessage] = React.useState<string>('');

  const requestId = React.useRef<number>(0);
  const mounted = React.useRef<boolean>(true);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const inputRef = React.useRef<HTMLInputElement>(null);

  const webUrl = context.pageContext.web.absoluteUrl;

  // --------------------------------------------------
  // Lifecycle helpers
  // --------------------------------------------------

  React.useEffect(() => {
    mounted.current = true;

    return () => {
      mounted.current = false;
    };
  }, []);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (
        containerRef.current &&
        !containerRef.current.contains(e.target as Node)
      ) {
        setOpen(false);
      }
    };

    document.addEventListener('mousedown', handler);

    return () => {
      document.removeEventListener('mousedown', handler);
    };
  }, []);

  // --------------------------------------------------
  // SharePoint calls
  // --------------------------------------------------

  // Search 1: tenant people search (the same one SharePoint's own picker uses).
  const searchDirectory = async (text: string): Promise<IPickedPerson[]> => {
    const types =
      props.principalTypes && props.principalTypes.length > 0
        ? props.principalTypes
        : [1];

    let principalType = 0;
    types.forEach((t) => {
      principalType = principalType | Number(t);
    });

    if (!principalType) {
      principalType = 1;
    }

    const response = await context.spHttpClient.post(
      `${webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface` +
      '.clientPeoplePickerSearchUser',
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          queryParams: {
            QueryString: text,
            MaximumEntitySuggestions: 10,
            AllowEmailAddresses: true,
            AllowOnlyEmailAddresses: false,
            AllowMultipleEntities: false,
            AllUrlZones: false,
            PrincipalSource: 15,
            PrincipalType: principalType,
            SharePointGroupID: 0
          }
        })
      }
    );

    if (!response.ok) {
      throw new Error(`${response.status}: ${await response.text()}`);
    }

    const data = await response.json() as IPickerSearchResponse;
    const raw =
      data.value ||
      (data.d && data.d.ClientPeoplePickerSearchUser) ||
      '[]';

    console.debug('People picker raw response:', raw);

    const entities = JSON.parse(raw) as IPickerEntity[];

    return entities
      .filter((entity) => !!entity.Key)
      .map((entity) => {
        const email =
          (entity.EntityData && entity.EntityData.Email) ||
          (entity.Description && entity.Description.indexOf('@') > -1
            ? entity.Description
            : '');

        const displayName = entity.DisplayText || entity.Key || '';

        return {
          loginName: entity.Key as string,
          displayName,
          email,
          text: displayName,
          secondaryText: email
        };
      });
  };

  const toPerson = (user: {
    Id?: number;
    Title?: string;
    Email?: string;
    LoginName?: string;
  }): IPickedPerson | undefined => {
    if (!user.LoginName) {
      return undefined;
    }

    const displayName = user.Title || user.LoginName;
    const email = user.Email || '';

    return {
      id: user.Id,
      loginName: user.LoginName,
      displayName,
      email,
      text: displayName,
      secondaryText: email
    };
  };

  // Search 2: people already known to this site (name, email or login).
  const searchSiteUsers = async (text: string): Promise<IPickedPerson[]> => {
    const q = encodeURIComponent(text.replace(/'/g, "''"));

    const url =
      `${webUrl}/_api/web/siteusers` +
      '?$select=Id,Title,Email,LoginName,PrincipalType' +
      `&$filter=PrincipalType eq 1 and (substringof('${q}',Title) or ` +
      `substringof('${q}',Email) or substringof('${q}',LoginName))` +
      '&$top=10';

    const response = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      throw new Error(`${response.status}: ${await response.text()}`);
    }

    const data = await response.json() as {
      value?: Array<{
        Id?: number;
        Title?: string;
        Email?: string;
        LoginName?: string;
      }>;
    };

    const people: IPickedPerson[] = [];

    (data.value || []).forEach((user) => {
      const person = toPerson(user);

      if (person) {
        people.push(person);
      }
    });

    return people;
  };

  // Search 3: a full email address typed in -> resolve it directly.
  const resolveByEmail = async (text: string): Promise<IPickedPerson[]> => {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text)) {
      return [];
    }

    const response = await context.spHttpClient.post(
      `${webUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({ logonName: `i:0#.f|membership|${text}` })
      }
    );

    if (!response.ok) {
      return [];
    }

    const person = toPerson(await response.json() as {
      Id?: number;
      Title?: string;
      Email?: string;
      LoginName?: string;
    });

    return person ? [person] : [];
  };

  // Tries each search in turn until one returns people.
  const searchUsers = async (text: string): Promise<IPickedPerson[]> => {
    const names = ['people search', 'site users', 'email lookup'];

    const attempts: Array<() => Promise<IPickedPerson[]>> = [
      () => searchDirectory(text),
      () => searchSiteUsers(text),
      () => resolveByEmail(text)
    ];

    let failures = 0;

    for (let i = 0; i < attempts.length; i++) {
      try {
        const results = await attempts[i]();

        console.debug(
          `People picker (${names[i]}): ${results.length} result(s) for "${text}"`
        );

        if (results.length > 0) {
          return results;
        }
      } catch (error) {
        failures++;
        console.warn(`People picker ${names[i]} failed:`, error);
      }
    }

    if (failures === attempts.length) {
      throw new Error('All people search methods failed.');
    }

    return [];
  };

  const ensureUserId = async (loginName: string): Promise<number | undefined> => {
    const response = await context.spHttpClient.post(
      `${webUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({ logonName: loginName })
      }
    );

    if (!response.ok) {
      throw new Error(`${response.status}: ${await response.text()}`);
    }

    const user = await response.json() as { Id?: number };

    return user.Id;
  };

  // --------------------------------------------------
  // Search while typing (debounced)
  // --------------------------------------------------

  React.useEffect(() => {
    const text = query.trim();

    if (text.length < 2) {
      requestId.current++;
      setSuggestions([]);
      setMessage('');
      setBusy(false);
      return undefined;
    }

    requestId.current++;
    const current = requestId.current;

    setBusy(true);

    const timer = window.setTimeout(() => {
      searchUsers(text)
        .then((results) => {
          if (!mounted.current || current !== requestId.current) {
            return;
          }

          const selected = value.map((p) => p.loginName.toLowerCase());
          const filtered = results.filter(
            (p) => selected.indexOf(p.loginName.toLowerCase()) === -1
          );

          setSuggestions(filtered);
          setMessage(
            filtered.length === 0
              ? 'No matching people found. Try the full email address.'
              : ''
          );
          setOpen(true);
        })
        .catch((error: unknown) => {
          if (!mounted.current || current !== requestId.current) {
            return;
          }

          console.error('People search failed:', error);
          setSuggestions([]);
          setMessage('Could not search for people. Check the browser console.');
          setOpen(true);
        })
        .then(() => {
          if (mounted.current && current === requestId.current) {
            setBusy(false);
          }
        }, () => undefined);
    }, props.resolveDelay !== undefined ? props.resolveDelay : 300);

    return () => {
      window.clearTimeout(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query, props.resolveDelay]);

  // --------------------------------------------------
  // Select / remove
  // --------------------------------------------------

  const select = async (person: IPickedPerson): Promise<void> => {
    if (value.length >= limit) {
      return;
    }

    let picked: IPickedPerson = person;

    if (props.ensureUser !== false) {
      try {
        const id = await ensureUserId(person.loginName);

        if (id) {
          picked = { ...person, id };
        }
      } catch (error) {
        // The submission service resolves the user again on save.
        console.warn('Could not resolve user ID now:', error);
      }
    }

    if (!mounted.current) {
      return;
    }

    onChange(value.concat([picked]));

    requestId.current++;
    setQuery('');
    setSuggestions([]);
    setMessage('');
    setBusy(false);
    setOpen(false);
  };

  const remove = (person: IPickedPerson): void => {
    onChange(value.filter((p) => p.loginName !== person.loginName));
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      // Never submit the form from the picker.
      e.preventDefault();

      if (suggestions.length > 0) {
        select(suggestions[0]).catch((error: unknown) => {
          console.error('Could not select person:', error);
        });
      }
    } else if (e.key === 'Escape') {
      setOpen(false);
    } else if (e.key === 'Backspace' && query === '' && value.length > 0) {
      remove(value[value.length - 1]);
    }
  };

  // --------------------------------------------------
  // Render
  // --------------------------------------------------

  const atLimit = value.length >= limit;
  const showList =
    open && !atLimit && (busy || message !== '' || suggestions.length > 0);

  return (
    <div ref={containerRef} style={{ position: 'relative' }}>
      <div
        style={boxStyle}
        onClick={() => {
          if (inputRef.current) {
            inputRef.current.focus();
          }
        }}
      >
        {value.map((person) => (
          <span
            key={person.loginName}
            style={
              props.showProfilePicture
                ? { ...chipStyle, paddingLeft: 2 }
                : chipStyle
            }
          >
            {props.showProfilePicture && (
              <ProfilePicture
                webUrl={webUrl}
                displayName={person.displayName}
                email={person.email}
                loginName={person.loginName}
                size={props.profilePictureSize}
              />
            )}
            {person.displayName}
            <button
              type="button"
              style={chipButtonStyle}
              aria-label={`Remove ${person.displayName}`}
              onClick={() => remove(person)}
            >
              ×
            </button>
          </span>
        ))}

        {!atLimit && (
          <input
            ref={inputRef}
            type="text"
            style={inputStyle}
            value={query}
            placeholder={
              value.length === 0
                ? props.placeholder || 'Search for a person'
                : ''
            }
            aria-label={props.ariaLabel || 'People picker'}
            autoComplete="off"
            onChange={(e) => {
              setQuery(e.currentTarget.value);
              setOpen(true);
            }}
            onFocus={() => setOpen(true)}
            onKeyDown={onKeyDown}
          />
        )}
      </div>

      {showList && (
        <ul role="listbox" style={listStyle}>
          {busy && suggestions.length === 0 && (
            <li style={infoStyle}>Searching...</li>
          )}

          {!busy && message !== '' && (
            <li style={infoStyle}>{message}</li>
          )}

          {suggestions.map((person) => (
            <li
              key={person.loginName}
              role="option"
              aria-selected={false}
              style={
                props.showProfilePicture
                  ? {
                      ...itemStyle,
                      display: 'flex',
                      alignItems: 'center',
                      gap: 10
                    }
                  : itemStyle
              }
              onMouseDown={(e) => {
                // Keep focus so the list does not close before the click.
                e.preventDefault();

                select(person).catch((error: unknown) => {
                  console.error('Could not select person:', error);
                });
              }}
            >
              {props.showProfilePicture && (
                <ProfilePicture
                  webUrl={webUrl}
                  displayName={person.displayName}
                  email={person.email}
                  loginName={person.loginName}
                  size={props.profilePictureSize}
                />
              )}

              <div>
                <div>{person.displayName}</div>

                {person.email && (
                  <div style={{ fontSize: 12, color: '#605e5c' }}>
                    {person.email}
                  </div>
                )}
              </div>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};

export default SharePointPeoplePicker;