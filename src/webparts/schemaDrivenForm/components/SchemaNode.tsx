// import * as React from 'react';
// import { IJsonSchema } from '../models/ISchemaDrivenFormProps';
// import styles from './SchemaDrivenForm.module.scss';

// export interface ISchemaNodeProps {
//   name: string;
//   schema: IJsonSchema;
//   rootSchema: IJsonSchema;
//   value: unknown;
//   required: boolean;
//   onChange: (name: string, value: unknown) => void;
//   path: string;
// }

// function localRef(root: IJsonSchema, ref: string): IJsonSchema | undefined {
//   if (ref.indexOf('#/') !== 0) return undefined;

//   let current: unknown = root;

//   const parts = ref.substring(2).split('/').map(
//     (p) => p.replace(/~1/g, '/').replace(/~0/g, '~')
//   );

//   for (const part of parts) {
//     if (!current || typeof current !== 'object') return undefined;
//     current = (current as { [key: string]: unknown })[part];
//   }

//   return current && typeof current === 'object'
//     ? current as IJsonSchema
//     : undefined;
// }

// function resolved(root: IJsonSchema, schema: IJsonSchema): IJsonSchema {
//   if (typeof schema.$ref !== 'string') return schema;

//   const target = localRef(root, schema.$ref);
//   if (!target) return schema;

//   const siblings = Object.assign({}, schema);
//   delete siblings.$ref;

//   return Object.assign({}, target, siblings);
// }

// function getType(schema: IJsonSchema): string {
//   const types = Array.isArray(schema.type)
//     ? schema.type
//     : [schema.type || (schema.properties ? 'object' : schema.items ? 'array' : 'string')];

//   return types.find((type) => type !== 'null') || 'string';
// }

// const SchemaNode: React.FC<ISchemaNodeProps> = (props) => {
//   const schema = resolved(props.rootSchema, props.schema);

//   const title = String(schema.title || props.name);
//   const description = typeof schema.description === 'string'
//     ? schema.description
//     : '';

//   const type = getType(schema);
//   const value = props.value;

//   const update = (next: unknown): void => {
//     props.onChange(props.path, next);
//   };

//   // =====================================================
//   // OBJECT: render ALL properties
//   // =====================================================

//   if (type === 'object' || schema.properties) {
//     const objectValue = (
//       value && typeof value === 'object' && !Array.isArray(value)
//         ? value
//         : {}
//     ) as { [key: string]: unknown };

//     const childProperties = schema.properties || {};
//     const keys = Object.keys(childProperties);

//     return (
//       <fieldset className={styles.objectGroup}>
//         <legend className={styles.objectLegend}>
//           {title}
//           {props.required ? ' *' : ''}
//         </legend>

//         {description && <p className={styles.help}>{description}</p>}

//         {keys.map((key) => {
//           const childSchema = childProperties[key];

//           const childValue = objectValue[key] !== undefined
//             ? objectValue[key]
//             : childSchema.default;

//           return (
//             <SchemaNode
//               key={`${props.path}.${key}`}
//               name={key}
//               path={props.path ? `${props.path}.${key}` : key}
//               schema={childSchema}
//               rootSchema={props.rootSchema}
//               value={childValue}
//               required={(schema.required || []).indexOf(key) >= 0}
//               onChange={(_childPath, nextChildValue) => {
//                 const next = Object.assign({}, objectValue);
//                 next[key] = nextChildValue;
//                 update(next);
//               }}
//             />
//           );
//         })}
//       </fieldset>
//     );
//   }

//   // =====================================================
//   // ARRAY: render ALL items
//   // =====================================================

//   if (type === 'array') {
//     const itemsSchema = (
//       !Array.isArray(schema.items) && schema.items
//         ? schema.items
//         : { type: 'string' }
//     ) as IJsonSchema;

//     const items = Array.isArray(value) ? value : [];

//     const createItem = (): unknown => {
//       if (itemsSchema.default !== undefined) {
//         return itemsSchema.default;
//       }

//       const itemType = getType(resolved(props.rootSchema, itemsSchema));

//       if (itemType === 'object') return {};
//       if (itemType === 'array') return [];
//       if (itemType === 'boolean') return false;

//       return '';
//     };

//     return (
//       <div className={styles.field}>
//         <label className={styles.label}>
//           {title}
//           {props.required ? ' *' : ''}
//         </label>

//         {description && <p className={styles.help}>{description}</p>}

//         {items.map((item, index) => (
//           <div
//             // className={styles.arrayItem}
//             key={`${props.path}-${index}`}
//           >
//             <SchemaNode
//               name={`${props.name} ${index + 1}`}
//               path={`${props.path}[${index}]`}
//               schema={itemsSchema}
//               rootSchema={props.rootSchema}
//               value={item}
//               required={false}
//               onChange={(_childPath, nextValue) => {
//                 const nextItems = items.slice();
//                 nextItems[index] = nextValue;
//                 update(nextItems);
//               }}
//             />

//             <button
//               type="button"
//               className={styles.removeButton}
//               onClick={() => update(items.filter((_item, i) => i !== index))}
//             >
//               Remove
//             </button>
//           </div>
//         ))}

//         <button
//           type="button"
//           className={styles.secondaryButton}
//           onClick={() => update(items.concat([createItem()]))}
//         >
//           {items.length === 0 ? 'Add first item' : 'Add another item'}
//         </button>
//       </div>
//     );
//   }

//   // =====================================================
//   // PRIMITIVE FIELDS
//   // =====================================================

//   const common = {
//     id: props.path,
//     name: props.path,
//     required: props.required,
//     'aria-describedby': description ? `${props.path}-help` : undefined
//   };

//   const examples = Array.isArray(schema.examples) ? schema.examples : [];

//   const placeholder = typeof examples[0] === 'string'
//     ? examples[0]
//     : `Enter ${title}`;

//   return (
//     <div className={styles.field}>
//       <label className={styles.label} htmlFor={props.path}>
//         {title}
//         {props.required ? ' *' : ''}
//       </label>

//       {description && (
//         <p className={styles.help} id={`${props.path}-help`}>
//           {description}
//         </p>
//       )}

//       {Array.isArray(schema.enum) ? (
//         <select
//           {...common}
//           className={styles.control}
//           value={value === undefined || value === null ? '' : String(value)}
//           onChange={(e) => update(
//             e.currentTarget.value === '' ? undefined : e.currentTarget.value
//           )}
//         >
//           <option value="">Select {title}</option>

//           {schema.enum.map((option, index) => (
//             <option key={index} value={String(option)}>
//               {String(option)}
//             </option>
//           ))}
//         </select>
//       ) : type === 'boolean' ? (
//         <label className={styles.checkboxLabel}>
//           <input
//             {...common}
//             type="checkbox"
//             checked={value === true}
//             onChange={(e) => update(e.currentTarget.checked)}
//           />
//           {title}
//         </label>
//       ) : (
//         <input
//           {...common}
//           className={styles.control}
//           type={
//             schema.format === 'date'
//               ? 'date'
//               : schema.format === 'date-time'
//                 ? 'datetime-local'
//                 : schema.format === 'email'
//                   ? 'email'
//                   : type === 'integer' || type === 'number'
//                     ? 'number'
//                     : 'text'
//           }
//           value={value === undefined || value === null ? '' : String(value)}
//           min={typeof schema.minimum === 'number' ? schema.minimum : undefined}
//           max={typeof schema.maximum === 'number' ? schema.maximum : undefined}
//           minLength={typeof schema.minLength === 'number' ? schema.minLength : undefined}
//           maxLength={typeof schema.maxLength === 'number' ? schema.maxLength : undefined}
//           pattern={typeof schema.pattern === 'string' ? schema.pattern : undefined}
//           placeholder={placeholder}
//           onChange={(e) => {
//             const raw = e.currentTarget.value;

//             if (type === 'integer') {
//               update(raw === '' ? undefined : parseInt(raw, 10));
//             } else if (type === 'number') {
//               update(raw === '' ? undefined : Number(raw));
//             } else {
//               update(raw);
//             }
//           }}
//         />
//       )}
//     </div>
//   );
// };

// export default SchemaNode;
import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IJsonSchema,
  IPeoplePickerConfig
} from '../models/ISchemaDrivenFormProps';
import { IFileInfo } from '../models/IFormDefinition';
import SharePointPeoplePicker, {
  IPickedPerson
} from '../services/SharePointPeoplePicker';
import styles from './SchemaDrivenForm.module.scss';

export interface ISchemaNodeProps {
  name: string;
  schema: IJsonSchema;
  rootSchema: IJsonSchema;
  value: unknown;
  required: boolean;
  onChange: (name: string, value: unknown) => void;
  path: string;
  /** Needed by the people picker. */
  context?: WebPartContext;
  /** Changes after a successful submit so file and people inputs reset. */
  resetKey?: number;
}

function localRef(root: IJsonSchema, ref: string): IJsonSchema | undefined {
  if (ref.indexOf('#/') !== 0) return undefined;

  let current: unknown = root;

  const parts = ref.substring(2).split('/').map(
    (p) => p.replace(/~1/g, '/').replace(/~0/g, '~')
  );

  for (const part of parts) {
    if (!current || typeof current !== 'object') return undefined;
    current = (current as { [key: string]: unknown })[part];
  }

  return current && typeof current === 'object'
    ? current as IJsonSchema
    : undefined;
}

function resolved(root: IJsonSchema, schema: IJsonSchema): IJsonSchema {
  if (typeof schema.$ref !== 'string') return schema;

  const target = localRef(root, schema.$ref);
  if (!target) return schema;

  const siblings = Object.assign({}, schema);
  delete siblings.$ref;

  return Object.assign({}, target, siblings);
}

function getType(schema: IJsonSchema): string {
  const types = Array.isArray(schema.type)
    ? schema.type
    : [schema.type || (schema.properties ? 'object' : schema.items ? 'array' : 'string')];

  return types.find((type) => type !== 'null') || 'string';
}

const SchemaNode: React.FC<ISchemaNodeProps> = (props) => {
  const schema = resolved(props.rootSchema, props.schema);

  const title = String(schema.title || props.name);
  const description = typeof schema.description === 'string'
    ? schema.description
    : '';

  const type = getType(schema);
  const value = props.value;
  const control = String(schema['x-control'] || '').toLowerCase();

  const update = (next: unknown): void => {
    props.onChange(props.path, next);
  };

  const requiredMark = props.required
    ? <span className={styles.required} aria-hidden="true">*</span>
    : null;

  const helpText = description
    ? <p className={styles.help} id={`${props.path}-help`}>{description}</p>
    : null;

  // =====================================================
  // PEOPLE PICKER  ("x-control": "peoplePicker")
  // =====================================================

  if (control === 'peoplepicker') {
    const config: IPeoplePickerConfig = schema['x-peoplePicker'] || {};
    const picked = Array.isArray(value) ? value as IPickedPerson[] : [];
    const multiple = config.allowMultiple === true || type === 'array';

    const limit = config.personSelectionLimit !== undefined
      ? config.personSelectionLimit
      : multiple
        ? (typeof schema.maxItems === 'number' ? schema.maxItems : 10)
        : 1;

    return (
      <div className={`${styles.field} ${styles.fieldFullWidth}`}>
        <label className={styles.label}>
          {title}
          {requiredMark}
        </label>

        {helpText}

        <div className={styles.peoplePicker}>
          {props.context ? (
            <SharePointPeoplePicker
              key={`${props.path}_${props.resetKey || 0}`}
              context={props.context}
              value={picked}
              limit={limit}
              principalTypes={
                config.principalTypes && config.principalTypes.length > 0
                  ? config.principalTypes
                  : [1]
              }
              showProfilePicture={config.showProfilePicture === true}
              profilePictureSize={config.profilePictureSize}
              placeholder={config.placeholder}
              ensureUser={config.ensureUser !== false}
              resolveDelay={config.resolveDelay}
              ariaLabel={title}
              onChange={(people) => update(people.length > 0 ? people : undefined)}
            />
          ) : (
            <p className={styles.help}>
              The people picker needs the web part context.
            </p>
          )}
        </div>
      </div>
    );
  }

  // =====================================================
  // FILE UPLOAD  ("x-control": "file" or "format": "binary")
  // =====================================================

  if (control === 'file' || schema.format === 'binary') {
    const files = Array.isArray(value) ? value as IFileInfo[] : [];
    const allowMultiple = type === 'array' ||
      (typeof schema.maxItems === 'number' && schema.maxItems > 1);

    return (
      <div className={`${styles.field} ${styles.fieldFullWidth}`}>
        <label className={styles.label} htmlFor={props.path}>
          {title}
          {requiredMark}
        </label>

        {helpText}

        <input
          key={`${props.path}_${props.resetKey || 0}`}
          id={props.path}
          type="file"
          multiple={allowMultiple}
          aria-describedby={description ? `${props.path}-help` : undefined}
          onChange={(e) => {
            const list = e.currentTarget.files;
            const infos: IFileInfo[] = [];

            if (list) {
              for (let i = 0; i < list.length; i++) {
                infos.push({
                  name: list[i].name,
                  size: list[i].size,
                  type: list[i].type,
                  file: list[i]
                });
              }
            }

            update(infos.length > 0 ? infos : undefined);
          }}
        />

        {files.length > 0 && (
          <ul className={styles.help}>
            {files.map((file, index) => (
              <li key={`${file.name}_${index}`}>
                {file.name} ({Math.round(file.size / 1024)} KB)
              </li>
            ))}
          </ul>
        )}
      </div>
    );
  }

  // =====================================================
  // OBJECT: render ALL properties
  // =====================================================

  if (type === 'object' || schema.properties) {
    const objectValue = (
      value && typeof value === 'object' && !Array.isArray(value)
        ? value
        : {}
    ) as { [key: string]: unknown };

    const childProperties = schema.properties || {};
    const keys = Object.keys(childProperties);

    return (
      <fieldset className={`${styles.objectGroup} ${styles.fieldFullWidth}`}>
        <legend className={styles.objectLegend}>
          {title}
          {requiredMark}
        </legend>

        {helpText}

        {keys.map((key) => {
          const childSchema = childProperties[key];

          const childValue = objectValue[key] !== undefined
            ? objectValue[key]
            : childSchema.default;

          return (
            <SchemaNode
              key={`${props.path}.${key}`}
              name={key}
              path={props.path ? `${props.path}.${key}` : key}
              schema={childSchema}
              rootSchema={props.rootSchema}
              value={childValue}
              required={(schema.required || []).indexOf(key) >= 0}
              context={props.context}
              resetKey={props.resetKey}
              onChange={(_childPath, nextChildValue) => {
                const next = Object.assign({}, objectValue);
                next[key] = nextChildValue;
                update(next);
              }}
            />
          );
        })}
      </fieldset>
    );
  }

  // =====================================================
  // ARRAY
  // =====================================================

  if (type === 'array') {
    const itemsSchema = (
      !Array.isArray(schema.items) && schema.items
        ? schema.items
        : { type: 'string' }
    ) as IJsonSchema;

    const resolvedItems = resolved(props.rootSchema, itemsSchema);
    const items = Array.isArray(value) ? value : [];

    // ----- array of fixed choices -> checkbox list -----
    if (Array.isArray(resolvedItems.enum)) {
      const selected = items.map((item) => String(item));

      return (
        <div className={`${styles.field} ${styles.fieldFullWidth}`}>
          <label className={styles.label}>
            {title}
            {requiredMark}
          </label>

          {helpText}

          {resolvedItems.enum.map((option, index) => {
            const text = String(option);

            return (
              <label key={`${text}_${index}`} className={styles.option}>
                <input
                  type="checkbox"
                  checked={selected.indexOf(text) >= 0}
                  onChange={(e) => {
                    const next = e.currentTarget.checked
                      ? items.concat([option])
                      : items.filter((item) => String(item) !== text);

                    update(next);
                  }}
                />
                {text}
              </label>
            );
          })}
        </div>
      );
    }

    // ----- array of anything else -> repeatable items -----
    const createItem = (): unknown => {
      if (resolvedItems.default !== undefined) {
        return resolvedItems.default;
      }

      const itemType = getType(resolvedItems);

      if (itemType === 'object') return {};
      if (itemType === 'array') return [];
      if (itemType === 'boolean') return false;

      return '';
    };

    const maxItems = typeof schema.maxItems === 'number' ? schema.maxItems : undefined;
    const canAdd = maxItems === undefined || items.length < maxItems;

    return (
      <div className={`${styles.field} ${styles.fieldFullWidth}`}>
        <label className={styles.label}>
          {title}
          {requiredMark}
        </label>

        {helpText}

        {items.map((item, index) => (
          <div
            className={styles.arrayItem}
            key={`${props.path}-${index}`}
          >
            <SchemaNode
              name={`${props.name} ${index + 1}`}
              path={`${props.path}[${index}]`}
              schema={itemsSchema}
              rootSchema={props.rootSchema}
              value={item}
              required={false}
              context={props.context}
              resetKey={props.resetKey}
              onChange={(_childPath, nextValue) => {
                const nextItems = items.slice();
                nextItems[index] = nextValue;
                update(nextItems);
              }}
            />

            <button
              type="button"
              className={styles.removeButton}
              onClick={() => update(items.filter((_item, i) => i !== index))}
            >
              Remove
            </button>
          </div>
        ))}

        {canAdd && (
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => update(items.concat([createItem()]))}
          >
            {items.length === 0 ? 'Add first item' : 'Add another item'}
          </button>
        )}
      </div>
    );
  }

  // =====================================================
  // PRIMITIVE FIELDS
  // =====================================================

  const common = {
    id: props.path,
    name: props.path,
    required: props.required,
    'aria-describedby': description ? `${props.path}-help` : undefined
  };

  const examples = Array.isArray(schema.examples) ? schema.examples : [];

  const placeholder = typeof examples[0] === 'string'
    ? examples[0]
    : `Enter ${title}`;

  // ----- boolean: the label sits next to the checkbox -----
  if (type === 'boolean') {
    return (
      <div className={styles.field}>
        <label className={styles.checkboxLabel} htmlFor={props.path}>
          <input
            {...common}
            type="checkbox"
            checked={value === true}
            onChange={(e) => update(e.currentTarget.checked)}
          />
          {title}
          {requiredMark}
        </label>

        {helpText}
      </div>
    );
  }

  const useTextarea = schema.format === 'textarea' ||
    (typeof schema.maxLength === 'number' && schema.maxLength > 200);

  return (
    <div className={useTextarea ? `${styles.field} ${styles.fieldFullWidth}` : styles.field}>
      <label className={styles.label} htmlFor={props.path}>
        {title}
        {requiredMark}
      </label>

      {helpText}

      {Array.isArray(schema.enum) ? (
        <select
          {...common}
          className={styles.control}
          value={value === undefined || value === null ? '' : String(value)}
          onChange={(e) => update(
            e.currentTarget.value === '' ? undefined : e.currentTarget.value
          )}
        >
          <option value="">Select {title}</option>

          {schema.enum.map((option, index) => (
            <option key={index} value={String(option)}>
              {String(option)}
            </option>
          ))}
        </select>
      ) : useTextarea ? (
        <textarea
          {...common}
          className={styles.control}
          rows={4}
          value={value === undefined || value === null ? '' : String(value)}
          minLength={typeof schema.minLength === 'number' ? schema.minLength : undefined}
          maxLength={typeof schema.maxLength === 'number' ? schema.maxLength : undefined}
          placeholder={placeholder}
          onChange={(e) => update(e.currentTarget.value)}
        />
      ) : (
        <input
          {...common}
          className={styles.control}
          type={
            schema.format === 'date'
              ? 'date'
              : schema.format === 'date-time'
                ? 'datetime-local'
                : schema.format === 'email'
                  ? 'email'
                  : type === 'integer' || type === 'number'
                    ? 'number'
                    : 'text'
          }
          value={value === undefined || value === null ? '' : String(value)}
          min={typeof schema.minimum === 'number' ? schema.minimum : undefined}
          max={typeof schema.maximum === 'number' ? schema.maximum : undefined}
          minLength={typeof schema.minLength === 'number' ? schema.minLength : undefined}
          maxLength={typeof schema.maxLength === 'number' ? schema.maxLength : undefined}
          pattern={typeof schema.pattern === 'string' ? schema.pattern : undefined}
          placeholder={placeholder}
          onChange={(e) => {
            const raw = e.currentTarget.value;

            if (type === 'integer') {
              update(raw === '' ? undefined : parseInt(raw, 10));
            } else if (type === 'number') {
              update(raw === '' ? undefined : Number(raw));
            } else {
              update(raw);
            }
          }}
        />
      )}
    </div>
  );
};

export default SchemaNode;