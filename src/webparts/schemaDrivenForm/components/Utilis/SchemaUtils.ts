import { IJsonSchema } from '../../models/ISchemaDrivenFormProps';

/* eslint-disable @typescript-eslint/no-explicit-any */

function isObj(v: any): boolean {
  return !!v && typeof v === 'object' && !Array.isArray(v);
}

function hasProps(s: any): boolean {
  return isObj(s) && isObj(s.properties) && Object.keys(s.properties).length > 0;
}

function resolveRef(root: any, schema: any): any {
  if (!isObj(schema) || typeof schema.$ref !== 'string' || schema.$ref.indexOf('#/') !== 0) {
    return schema;
  }

  let cur: any = root;
  const parts = schema.$ref.substring(2).split('/').map(
    (p: string) => p.replace(/~1/g, '/').replace(/~0/g, '~')
  );

  for (const part of parts) {
    cur = isObj(cur) ? cur[part] : undefined;
  }

  if (!isObj(cur)) return schema;

  const siblings = Object.assign({}, schema);
  delete siblings.$ref;
  return Object.assign({}, cur, siblings);
}

function mergeAllOf(root: any, schema: any): any {
  if (!isObj(schema) || !Array.isArray(schema.allOf)) return schema;

  const merged: any = Object.assign({}, schema);
  delete merged.allOf;
  merged.properties = Object.assign({}, merged.properties);
  merged.required = (merged.required || []).slice();

  schema.allOf.forEach((part: any) => {
    const p = mergeAllOf(root, resolveRef(root, part));
    if (!isObj(p)) return;
    Object.assign(merged.properties, p.properties || {});
    merged.required = merged.required.concat(p.required || []);
    if (!merged.title && p.title) merged.title = p.title;
    if (!merged.description && p.description) merged.description = p.description;
  });

  return merged;
}

// Resolves $ref, allOf and oneOf/anyOf (first option) at one level.
function expand(root: any, schema: any): any {
  let s: any = resolveRef(root, schema);
  s = mergeAllOf(root, s);

  if (isObj(s) && !hasProps(s)) {
    const alt = (s.oneOf || s.anyOf || [])[0];
    if (alt) {
      const a = mergeAllOf(root, resolveRef(root, alt));
      s = Object.assign({}, s, a);
      delete s.oneOf;
      delete s.anyOf;
    }
  }
  return s;
}

// Looks inside nested keys (schema wrapper, definitions, $defs, ...) for a schema with properties.
function findNested(root: any, node: any, depth: number): any {
  if (!isObj(node) || depth > 3) return undefined;

  const keys = Object.keys(node);
  for (const key of keys) {
    const child = node[key];
    if (!isObj(child)) continue;

    const expanded = expand(root, child);
    if (hasProps(expanded)) return expanded;

    const deeper = findNested(root, child, depth + 1);
    if (deeper) return deeper;
  }
  return undefined;
}

// Builds a schema from a sample value (used when the file is plain data, not a schema).
function infer(value: any, name: string): any {
  if (Array.isArray(value)) {
    return {
      type: 'array',
      title: name,
      items: value.length > 0 ? infer(value[0], name) : { type: 'string' }
    };
  }
  if (isObj(value)) {
    const properties: { [key: string]: any } = {};
    Object.keys(value).forEach((k) => {
      if (k.charAt(0) !== '$') properties[k] = infer(value[k], k);
    });
    return { type: 'object', title: name, properties };
  }
  if (typeof value === 'number') {
    return { type: value % 1 === 0 ? 'integer' : 'number', title: name };
  }
  if (typeof value === 'boolean') return { type: 'boolean', title: name };
  return { type: 'string', title: name };
}

/**
 * Returns an object schema with "properties" from whatever the file contains:
 *  - a normal JSON Schema (root $ref / allOf / oneOf / anyOf)
 *  - a schema nested under a wrapper key or definitions
 *  - a plain sample JSON object (fields are inferred from its keys)
 */
export function getRootSchema(schema: IJsonSchema): IJsonSchema {
  const original: any = schema;
  let root: any = expand(original, original);

  if (!hasProps(root)) {
    const nested = findNested(original, root, 0);
    if (nested) root = nested;
  }

  if (!hasProps(root) && isObj(original)) {
    const inferred = infer(original, 'Form');
    if (hasProps(inferred)) root = inferred;
  }

  if (isObj(root) && root.properties && !root.type) {
    root = Object.assign({}, root, { type: 'object' });
  }

  return root as IJsonSchema;
}