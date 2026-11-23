import * as React from 'react';
import {
  ISchemaDrivenFormProps,
  IPeoplePickerConfig
} from '../models/ISchemaDrivenFormProps';
import {
  IFormDefinition,
  IFormField,
  IFormNode,
  IFieldValues,
  IFileInfo,
  IDateRange,
  IFormAction,
  IFormStyles,
  IStyleMap
} from '../models/IFormDefinition';
import { SharePointSchemaReader } from '../services/SharePointSchemaReader';
import {
  SubmissionService,
  ISubmissionResult
} from '../services/SubmissionService';
import {
  parseDefinition,
  getInitialValues,
  isNodeVisible,
  collectValues,
  validateForm
} from '../components/Utilis/Formfull';
import SharePointPeoplePicker, {
  IPickedPerson
} from '../services/SharePointPeoplePicker';
import styles from './SchemaDrivenForm.module.scss';

/**
 * Styling now comes from SchemaDrivenForm.module.scss (class names).
 * The JSON schema ("style", "styles", "labelStyle", ...) is still honoured,
 * but only as an OPTIONAL inline override on top of the SCSS classes.
 * If the schema has no style information, the SCSS template is used as-is.
 */

// Field types that should span both columns of the .formGrid
const FULL_WIDTH_TYPES = [
  'textarea',
  'radio',
  'multi_select',
  'checklist',
  'date_range',
  'rating',
  'file',
  'list',
  'peoplepicker',
  'peoplePicker'
];

/** Joins class names, ignoring empty values. */
const cx = (...names: Array<string | false | undefined>): string =>
  names.filter((n) => !!n).join(' ');

const SchemaDrivenForm: React.FC<ISchemaDrivenFormProps> = (props) => {
  const [definition, setDefinition] = React.useState<IFormDefinition | undefined>(undefined);
  const [values, setValues] = React.useState<IFieldValues>({});
  const [loading, setLoading] = React.useState<boolean>(false);
  const [saving, setSaving] = React.useState<boolean>(false);
  const [loadError, setLoadError] = React.useState<string>('');
  const [errors, setErrors] = React.useState<string[]>([]);
  const [saveError, setSaveError] = React.useState<string>('');
  const [result, setResult] = React.useState<ISubmissionResult | undefined>(undefined);
  const [resultStatus, setResultStatus] = React.useState<string>('');
  const [resetKey, setResetKey] = React.useState<number>(0);

  // --------------------------------------------------
  // OPTIONAL SCHEMA STYLE OVERRIDES
  // --------------------------------------------------

  const sheet: IFormStyles = (definition && definition.styles) || {};

  /**
   * Inline overrides from the schema for a named part + any extra style maps.
   * Returns undefined when there is nothing to override, so the SCSS class wins.
   */
  const sx = (
    part: keyof IFormStyles,
    ...extra: Array<IStyleMap | undefined>
  ): React.CSSProperties | undefined => {
    const merged = Object.assign({}, sheet[part], ...extra) as React.CSSProperties;
    return Object.keys(merged).length > 0 ? merged : undefined;
  };

  // --------------------------------------------------
  // LOAD JSON SCHEMA
  // --------------------------------------------------

  React.useEffect(() => {
    let cancelled = false;

    setDefinition(undefined);
    setValues({});
    setLoadError('');
    setErrors([]);
    setSaveError('');
    setResult(undefined);

    if (!props.schemaFileServerRelativeUrl) {
      setLoading(false);
      return undefined;
    }

    setLoading(true);

    const schemaReader = new SharePointSchemaReader(props.context);

    schemaReader
      .read(props.schemaFileServerRelativeUrl)
      .then((raw) => {
        if (cancelled) {
          return;
        }

        const def = parseDefinition(raw);

        if (!def.fields || def.fields.length === 0) {
          const keys = Object.keys(raw as unknown as object).join(', ') || '(none)';

          setLoadError(
            `No fields were found in this file. Top-level keys: ${keys}`
          );
        } else {
          setDefinition(def);
          setValues(getInitialValues(def));
        }
      })
      .catch((e: unknown) => {
        if (cancelled) {
          return;
        }

        setLoadError(
          e instanceof Error ? e.message : String(e)
        );
      })
      .then(() => {
        if (!cancelled) {
          setLoading(false);
        }
      }, () => undefined);

    return () => {
      cancelled = true;
    };
  }, [props.context, props.schemaFileServerRelativeUrl]);

  // --------------------------------------------------
  // UPDATE FIELD VALUE
  // --------------------------------------------------

  const setValue = (id: string, value: unknown): void => {
    setValues((previous) => {
      const next = Object.assign({}, previous);

      if (value === undefined) {
        delete next[id];
      } else {
        next[id] = value;
      }

      return next;
    });
  };

  // --------------------------------------------------
  // SUBMIT TO SHAREPOINT LIST
  // --------------------------------------------------

  const submit = async (
    action: 'draft' | 'submitted'
  ): Promise<void> => {
    if (!definition || saving) {
      return;
    }

    setSaveError('');
    setResult(undefined);

    if (!props.submissionLibraryId) {
      setSaveError(
        'Select a SharePoint Submission List in the property pane.'
      );
      return;
    }

    if (action === 'submitted') {
      const problems = validateForm(definition, values);

      setErrors(problems);

      if (problems.length > 0) {
        return;
      }
    } else {
      setErrors([]);
    }

    setSaving(true);

    try {
      const submissionValues = collectValues(
        definition,
        values
      );

      const submissionService = new SubmissionService(
        props.context
      );

      const saved = await submissionService.save(
        props.submissionLibraryId,
        props.schemaFileServerRelativeUrl,
        definition.formName,
        submissionValues,
        action
      );

      setResult(saved);
      setResultStatus(action);

      if (action === 'submitted') {
        setValues(getInitialValues(definition));
        setResetKey((previous) => previous + 1);
      }
    } catch (e: unknown) {
      setSaveError(
        e instanceof Error ? e.message : String(e)
      );
    } finally {
      setSaving(false);
    }
  };

  // --------------------------------------------------
  // FIELD CONTROLS
  // --------------------------------------------------

  const renderControl = (
    f: IFormField
  ): React.ReactNode => {
    const id = f.fieldId;
    const value = values[id];
    const options = f.options || [];
    const controlStyle = sx('control', f.controlStyle);

    switch (f.type) {
      case 'textarea':
        return (
          <textarea
            id={id}
            className={styles.control}
            style={controlStyle}
            rows={4}
            placeholder={f.placeholder}
            value={typeof value === 'string' ? value : ''}
            onChange={(e) => setValue(id, e.currentTarget.value)}
          />
        );

      case 'date':
        return (
          <input
            id={id}
            type="date"
            className={styles.control}
            style={controlStyle}
            value={typeof value === 'string' ? value : ''}
            onChange={(e) => setValue(id, e.currentTarget.value)}
          />
        );

      case 'date_range': {
        const range = (value || {}) as IDateRange;

        return (
          <div className={styles.dateRange} style={sx('dateRange')}>
            <input
              type="date"
              className={styles.control}
              style={controlStyle}
              aria-label={`${f.label} start`}
              value={range.start || ''}
              onChange={(e) =>
                setValue(
                  id,
                  Object.assign({}, range, {
                    start: e.currentTarget.value
                  })
                )
              }
            />

            <span>to</span>

            <input
              type="date"
              className={styles.control}
              style={controlStyle}
              aria-label={`${f.label} end`}
              value={range.end || ''}
              onChange={(e) =>
                setValue(
                  id,
                  Object.assign({}, range, {
                    end: e.currentTarget.value
                  })
                )
              }
            />
          </div>
        );
      }

      case 'number':
        return (
          <input
            id={id}
            type="number"
            className={styles.control}
            style={controlStyle}
            placeholder={f.placeholder}
            min={
              f.validation && f.validation.min !== undefined
                ? f.validation.min
                : f.min
            }
            max={
              f.validation && f.validation.max !== undefined
                ? f.validation.max
                : f.max
            }
            value={typeof value === 'number' ? String(value) : ''}
            onChange={(e) => {
              const raw = e.currentTarget.value;
              setValue(id, raw === '' ? undefined : Number(raw));
            }}
          />
        );

      case 'radio':
        return (
          <div>
            {options.map((option) => (
              <label
                key={option.value}
                className={styles.option}
                style={sx('option')}
              >
                <input
                  type="radio"
                  name={id}
                  value={option.value}
                  checked={value === option.value}
                  onChange={() => setValue(id, option.value)}
                />
                {option.label}
              </label>
            ))}
          </div>
        );

      case 'dropdown':
        return (
          <select
            id={id}
            className={styles.control}
            style={controlStyle}
            value={typeof value === 'string' ? value : ''}
            onChange={(e) =>
              setValue(
                id,
                e.currentTarget.value === ''
                  ? undefined
                  : e.currentTarget.value
              )
            }
          >
            <option value="">Select {f.label}</option>
            {options.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        );

      case 'multi_select':
        return (
          <select
            id={id}
            multiple
            size={Math.min(Math.max(options.length, 3), 8)}
            className={styles.control}
            style={controlStyle}
            value={Array.isArray(value) ? value as string[] : []}
            onChange={(e) => {
              const selected: string[] = [];
              const opts = e.currentTarget.options;

              for (let i = 0; i < opts.length; i++) {
                if (opts[i].selected) {
                  selected.push(opts[i].value);
                }
              }

              setValue(id, selected);
            }}
          >
            {options.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        );

      case 'checklist': {
        const current = Array.isArray(value) ? value as string[] : [];

        return (
          <div>
            {options.map((option) => (
              <label
                key={option.value}
                className={styles.option}
                style={sx('option')}
              >
                <input
                  type="checkbox"
                  checked={current.indexOf(option.value) >= 0}
                  onChange={(e) => {
                    const next = e.currentTarget.checked
                      ? current.concat([option.value])
                      : current.filter((item) => item !== option.value);

                    setValue(id, next);
                  }}
                />
                {option.label}
              </label>
            ))}
          </div>
        );
      }

      case 'checkbox':
        return (
          <label
            className={styles.checkboxLabel}
            style={sx('option', f.labelStyle)}
          >
            <input
              id={id}
              type="checkbox"
              checked={value === true}
              onChange={(e) => setValue(id, e.currentTarget.checked)}
            />
            {f.label}
            {f.required && <span className={styles.required}>*</span>}
          </label>
        );

      case 'rating': {
        const min = f.min !== undefined ? f.min : 1;
        const max = f.max !== undefined ? f.max : 5;
        const numbers: number[] = [];

        for (let i = min; i <= max; i++) {
          numbers.push(i);
        }

        return (
          <div className={styles.rating} style={sx('rating')}>
            {numbers.map((number) => {
              const selected = value === number;

              return (
                <button
                  key={number}
                  type="button"
                  className={selected ? styles.ratingSelected : styles.ratingButton}
                  style={
                    selected
                      ? Object.assign({}, sx('ratingButton'), sx('ratingSelected'))
                      : sx('ratingButton')
                  }
                  aria-pressed={selected}
                  onClick={() =>
                    setValue(id, selected ? undefined : number)
                  }
                >
                  {number}
                </button>
              );
            })}
          </div>
        );
      }

      case 'file': {
        const rules = f.validation || {};
        const files = Array.isArray(value) ? value as IFileInfo[] : [];

        return (
          <div>
            {/* Styled by ".field input[type='file']" in the SCSS */}
            <input
              key={`${id}_${resetKey}`}
              id={id}
              type="file"
              style={f.controlStyle as React.CSSProperties | undefined}
              multiple={(rules.maxFiles || 1) > 1}
              accept={
                rules.allowedExtensions
                  ? rules.allowedExtensions.join(',')
                  : undefined
              }
              onChange={(e) => {
                const list = e.currentTarget.files;
                const infos: IFileInfo[] = [];

                if (list) {
                  for (let i = 0; i < list.length; i++) {
                    // Keep the real File object so it can be uploaded
                    // as a SharePoint list item attachment.
                    infos.push({
                      name: list[i].name,
                      size: list[i].size,
                      type: list[i].type,
                      file: list[i]
                    } as IFileInfo);
                  }
                }

                setValue(id, infos.length > 0 ? infos : undefined);
              }}
            />

            {files.length > 0 && (
              <ul className={styles.help} style={sx('help', f.helpStyle)}>
                {files.map((file, index) => (
                  <li key={index}>
                    {file.name} ({Math.round(file.size / 1024)} KB)
                  </li>
                ))}
              </ul>
            )}
          </div>
        );
      }

      case 'list':
        return (
          <input
            id={id}
            type="text"
            className={styles.control}
            style={controlStyle}
            placeholder={f.placeholder || 'Separate values with commas'}
            value={
              typeof value === 'string'
                ? value
                : Array.isArray(value)
                  ? (value as string[]).join(', ')
                  : ''
            }
            onChange={(e) => setValue(id, e.currentTarget.value)}
          />
        );

      case 'peoplepicker':
      case 'peoplePicker': {
        const config = (f.peoplePicker || {}) as unknown as IPeoplePickerConfig;
        const picked = Array.isArray(value) ? value as IPickedPerson[] : [];

        const limit =
          config.personSelectionLimit !== undefined
            ? config.personSelectionLimit
            : config.allowMultiple
              ? 10
              : 1;

        const principalTypes =
          config.principalTypes && config.principalTypes.length > 0
            ? config.principalTypes.map((type) => Number(type))
            : [1];

        return (
          <div className={styles.peoplePicker}>
            <SharePointPeoplePicker
              key={`${id}_${resetKey}`}
              context={props.context}
              value={picked}
              limit={limit}
              principalTypes={principalTypes}
              showProfilePicture={config.showProfilePicture === true}
              profilePictureSize={config.profilePictureSize}
              placeholder={config.placeholder || f.placeholder}
              ensureUser={config.ensureUser !== false}
              resolveDelay={config.resolveDelay}
              ariaLabel={f.label}
              onChange={(people) =>
                setValue(id, people.length > 0 ? people : undefined)
              }
            />
          </div>
        );
      }

      default:
        return (
          <input
            id={id}
            type="text"
            className={styles.control}
            style={controlStyle}
            placeholder={f.placeholder}
            value={typeof value === 'string' ? value : ''}
            onChange={(e) => setValue(id, e.currentTarget.value)}
          />
        );
    }
  };

  // --------------------------------------------------
  // RENDER LOADING AND ERROR STATES
  // --------------------------------------------------

  if (!props.schemaFileServerRelativeUrl) {
    return (
      <div className={styles.container}>
        <div className={styles.emptyState}>
          Select a Schema Document Library and JSON Schema File
          in the property pane.
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className={styles.container}>
        <div className={styles.loading}>Loading form...</div>
      </div>
    );
  }

  if (loadError || !definition) {
    return (
      <div className={styles.container}>
        <div className={styles.error}>
          <strong>Form not loaded</strong>
          <p>{loadError}</p>
        </div>

        <p className={styles.help}>
          Check the Property Pane configuration and confirm
          that the web part has permission to read the JSON file.
        </p>
      </div>
    );
  }

  // --------------------------------------------------
  // RECURSIVE NESTED SCHEMA RENDERING
  // --------------------------------------------------

  const renderNode = (node: IFormNode, topLevel: boolean): React.ReactNode => {
    if (!isNodeVisible(definition, node.id, values)) {
      return null;
    }

    if (node.kind === 'group') {
      return (
        <fieldset
          key={node.id}
          className={cx(styles.objectGroup, topLevel && styles.fieldFullWidth)}
          style={sx('group', node.style)}
        >
          <legend
            className={styles.objectLegend}
            style={sx('legend', node.legendStyle)}
          >
            {node.title}
          </legend>

          {node.description && (
            <p className={styles.help} style={sx('help')}>
              {node.description}
            </p>
          )}

          {node.children.map((child) => renderNode(child, false))}
        </fieldset>
      );
    }

    const field = node.field;
    const fullWidth = FULL_WIDTH_TYPES.indexOf(field.type) >= 0;

    return (
      <div
        key={node.id}
        className={cx(styles.field, topLevel && fullWidth && styles.fieldFullWidth)}
        style={sx('field', field.style)}
      >
        {field.type !== 'checkbox' && (
          <label
            className={styles.label}
            style={sx('label', field.labelStyle)}
            htmlFor={field.fieldId}
          >
            {field.label}
            {field.required && <span className={styles.required}>*</span>}
          </label>
        )}

        {field.helpText && (
          <p className={styles.help} style={sx('help', field.helpStyle)}>
            {field.helpText}
          </p>
        )}

        {renderControl(field)}
      </div>
    );
  };

  // --------------------------------------------------
  // FORM ACTIONS
  // --------------------------------------------------

  const actions: IFormAction[] =
    definition.actions.length > 0
      ? definition.actions
      : [
          {
            actionId: 'submit',
            label: 'Submit',
            type: 'submit'
          }
        ];

  // --------------------------------------------------
  // MAIN FORM UI
  // --------------------------------------------------

  return (
    <div className={styles.container} style={sx('container', definition.style)}>
      <h2 className={styles.title} style={sx('title')}>
        {definition.formName}
      </h2>

      <form
        className={styles.form}
        onSubmit={(e) => {
          e.preventDefault();

          submit('submitted').catch((error: unknown) => {
            console.error('Form submission error:', error);
          });
        }}
        noValidate
      >
        <div className={styles.formGrid}>
          {definition.nodes.map((node) => renderNode(node, true))}
        </div>

        {errors.length > 0 && (
          <div className={styles.error} style={sx('error')}>
            <strong>Please fix the following:</strong>

            <ul>
              {errors.map((message, index) => (
                <li key={index}>{message}</li>
              ))}
            </ul>
          </div>
        )}

        {saveError && (
          <div className={styles.error} style={sx('error')}>
            {saveError}
          </div>
        )}

        {result && (
          <div className={styles.success} style={sx('success')}>
            <p>
              {resultStatus === 'draft'
                ? 'Draft saved successfully.'
                : 'Form submitted successfully.'}
            </p>

            <p>
              Reference: <strong>{result.reference}</strong>
            </p>

            <p>
              SharePoint List Item ID: <strong>{result.itemId}</strong>
            </p>

            {result.attachmentsUploaded > 0 && (
              <p>
                Attachments uploaded:{' '}
                <strong>{result.attachmentsUploaded}</strong>
              </p>
            )}
{result.emailSent ? (
  <p style={{ color: 'green' }}>
    Email request accepted by Microsoft Graph.
    Recipient: karthiraj@infinitiresearch.com
  </p>
) : result.emailError ? (
  <p style={{ color: 'red' }}>
    Email sending failed: {result.emailError}
  </p>
) : resultStatus === 'submitted' ? (
  <p style={{ color: 'red' }}>
    Email was not sent. Please check the browser console.
  </p>
) : null}
            {result.failedAttachments.length > 0 && (
              <p>
                The item was saved, but these files could not be attached:{' '}
                {result.failedAttachments.join(', ')}
              </p>
            )}

            {result.itemUrl && (
              <p>
                <a
                  href={result.itemUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  View submitted item
                </a>
              </p>
            )}
          </div>
        )}

        <div className={styles.actions} style={sx('actions')}>
          {actions.map((action) => {
            const isSecondary = action.type === 'save';

            return (
              <button
                key={action.actionId}
                type="button"
                className={isSecondary ? styles.secondaryButton : styles.primaryButton}
                style={
                  isSecondary
                    ? sx('secondaryButton', action.style)
                    : sx('primaryButton', action.style)
                }
                disabled={saving}
                onClick={() => {
                  submit(
                    isSecondary ? 'draft' : 'submitted'
                  ).catch((error: unknown) => {
                    console.error('Form submission error:', error);
                  });
                }}
              >
                {saving ? 'Please wait...' : action.label}
              </button>
            );
          })}
        </div>
      </form>
    </div>
  );
};

export default SchemaDrivenForm;
//karthigjhgjv

//jusajdvjasbda
