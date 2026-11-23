
import {
  IJsonSchema,
  ISchemaPropertyInfo,
  IPeoplePickerConfig
} from '../models/ISchemaDrivenFormProps';

function resolveLocalRef(
  root: IJsonSchema,
  ref: string
): IJsonSchema | undefined {

  if (!ref || ref.indexOf('#/') !== 0) {
    return undefined;
  }

  const parts = ref
    .substring(2)
    .split('/')
    .map((part) =>
      part.replace(/~1/g, '/').replace(/~0/g, '~')
    );

  let current: unknown = root;

  for (const part of parts) {
    if (!current || typeof current !== 'object') {
      return undefined;
    }

    current = (current as { [key: string]: unknown })[part];
  }

  return current &&
    typeof current === 'object' &&
    !Array.isArray(current)
    ? current as IJsonSchema
    : undefined;
}

function dereference(
  root: IJsonSchema,
  schema: IJsonSchema,
  depth: number = 0
): IJsonSchema {

  if (depth > 20 || typeof schema.$ref !== 'string') {
    return schema;
  }

  const target = resolveLocalRef(root, schema.$ref);

  if (!target) {
    return schema;
  }

  const siblings = Object.assign({}, schema);
  delete siblings.$ref;

  return dereference(
    root,
    Object.assign({}, target, siblings),
    depth + 1
  );
}

function inferType(schema: IJsonSchema): string {

  if (Array.isArray(schema.type)) {
    return schema.type.join(' | ');
  }

  if (typeof schema.type === 'string') {
    return schema.type;
  }

  if (schema.enum) {
    return 'enum';
  }

  if (schema.properties) {
    return 'object';
  }

  if (schema.items) {
    return 'array';
  }

  return 'unspecified';
}

function getPeoplePickerConfig(
  schema: IJsonSchema
): IPeoplePickerConfig | undefined {

  if (schema['x-control'] !== 'peoplePicker') {
    return undefined;
  }

  const config = schema['x-peoplePicker'];

  if (!config || typeof config !== 'object' || Array.isArray(config)) {
    return {};
  }

  return config as IPeoplePickerConfig;
}

export class SchemaInspector {

  public static inspect(
    root: IJsonSchema
  ): ISchemaPropertyInfo[] {

    const output: ISchemaPropertyInfo[] = [];

    const visit = (
      schemaInput: IJsonSchema,
      parentPath: string,
      requiredNames: string[],
      depth: number
    ): void => {

      if (depth > 20) {
        return;
      }

      const schema = dereference(root, schemaInput);

      Object.keys(schema.properties || {}).forEach((name) => {

        const child = (schema.properties || {})[name];

        const resolvedChild = dereference(root, child);

        const path = parentPath
          ? `${parentPath}.${name}`
          : name;

        const peoplePicker = getPeoplePickerConfig(child)
          || getPeoplePickerConfig(resolvedChild);

        output.push({
          path,
          title: String(child.title || resolvedChild.title || name),
          type: inferType(resolvedChild),
          required: requiredNames.indexOf(name) >= 0,
          format: typeof resolvedChild.format === 'string'
            ? resolvedChild.format
            : undefined,
          description: typeof resolvedChild.description === 'string'
            ? resolvedChild.description
            : undefined,

          // Dynamically identify custom controls
          control: typeof child['x-control'] === 'string'
            ? child['x-control'] as string
            : typeof resolvedChild['x-control'] === 'string'
              ? resolvedChild['x-control'] as string
              : undefined,

          // People Picker settings from JSON
          peoplePicker
        });

        // Nested object properties
        if (resolvedChild.properties) {
          visit(
            resolvedChild,
            path,
            resolvedChild.required || [],
            depth + 1
          );
        }

        // Nested array object properties
        if (
          resolvedChild.items &&
          !Array.isArray(resolvedChild.items)
        ) {

          const itemSchema = dereference(
            root,
            resolvedChild.items as IJsonSchema
          );

          if (itemSchema.properties) {
            visit(
              itemSchema,
              `${path}[]`,
              itemSchema.required || [],
              depth + 1
            );
          }
        }
      });
    };

    visit(
      root,
      '',
      root.required || [],
      0
    );

    return output;
  }
}
