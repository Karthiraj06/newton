import Ajv from 'ajv';
import { IJsonSchema } from '../models/ISchemaDrivenFormProps';

export interface IValidationResult {
  valid: boolean;
  messages: string[];
}

interface IAjvError {
  instancePath?: string; // Ajv v8
  dataPath?: string;     // Ajv v6
  message?: string;
}

interface ISchemaLike {
  type?: string | string[];
  properties?: { [key: string]: ISchemaLike };
}

export class SchemaValidationService {
  /**
   * Cleans the raw form values before validation / saving:
   *  - converts numeric strings to numbers for integer / number fields
   *  - converts "true"/"false" strings to booleans
   *  - drops empty strings and empty objects (so optional blocks that were
   *    never filled in do not fail validation with "required property ...")
   */
  public static normalize(schema: IJsonSchema, values: { [key: string]: unknown }): { [key: string]: unknown } {
    return SchemaValidationService.normalizeObject(schema as unknown as ISchemaLike, values);
  }

  private static normalizeObject(schema: ISchemaLike, value: { [key: string]: unknown }): { [key: string]: unknown } {
    const result: { [key: string]: unknown } = {};
    Object.keys(value || {}).forEach((key) => {
      const propSchema = schema.properties ? schema.properties[key] : undefined;
      const raw = value[key];
      const type = propSchema && !Array.isArray(propSchema.type) ? propSchema.type : undefined;

      if (raw === undefined || raw === null) return;

      if (type === 'object' && typeof raw === 'object' && propSchema) {
        const nested = SchemaValidationService.normalizeObject(propSchema, raw as { [key: string]: unknown });
        if (Object.keys(nested).length > 0) result[key] = nested;
        return;
      }

      if (type === 'integer' || type === 'number') {
        if (typeof raw === 'string') {
          if (raw.trim() === '') return;
          const num = Number(raw);
          if (isFinite(num)) result[key] = num;
          else result[key] = raw; // let validation report the type error
        } else {
          result[key] = raw;
        }
        return;
      }

      if (type === 'boolean') {
        if (raw === 'true') result[key] = true;
        else if (raw === 'false') result[key] = false;
        else result[key] = raw;
        return;
      }

      if (typeof raw === 'string' && raw.trim() === '') return;

      result[key] = raw;
    });
    return result;
  }

  public validate(schema: IJsonSchema, values: { [key: string]: unknown }): IValidationResult {
    // Ajv only knows draft-07 by default, so remove "$schema" (e.g. draft 2020-12).
    const schemaCopy = Object.assign({}, schema) as { [key: string]: unknown };
    delete schemaCopy.$schema;

    const options = { allErrors: true, strict: false, validateFormats: false };
    const ajv = new Ajv(options as never);

    const validateFn = ajv.compile(schemaCopy);
    const valid = validateFn(values) as boolean;

    if (valid) {
      return { valid: true, messages: [] };
    }

    const errors = (validateFn.errors || []) as IAjvError[];
    const messages = errors.map((e) => {
      const rawPath = e.instancePath || e.dataPath || '';
      const path = rawPath.replace(/^[./]/, '').replace(/\//g, '.');
      return path ? `${path}: ${e.message}` : `${e.message}`;
    });

    return { valid: false, messages };
  }
}