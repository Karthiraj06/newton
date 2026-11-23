import {
  IFormDefinition, IFormNode, IFormGroup, IFormField, IFormAction, IVisibleWhen,
  IFieldOption, IFieldValues, IFileInfo, IDateRange, IIndexEntry, IPeoplePickerConfig,
  IStyleMap, IFormStyles
} from '../../models/IFormDefinition';

/* eslint-disable @typescript-eslint/no-explicit-any */

interface IBuildCtx {
  fields: IFormField[];
  index: { [id: string]: IIndexEntry };
}

const CONTAINER_KEYS = ['fields', 'questions', 'children', 'sections', 'subSections', 'groups'];
const GROUP_TYPES = ['group', 'section', 'object', 'fieldset', 'panel'];

const STYLE_PARTS = [
  'container', 'title', 'group', 'legend', 'field', 'label', 'help', 'control',
  'option', 'dateRange', 'rating', 'ratingButton', 'ratingSelected', 'actions',
  'primaryButton', 'secondaryButton', 'error', 'success'
];

const TYPE_ALIASES: { [key: string]: string } = {
  string: 'text', input: 'text', email: 'text', phone: 'text', url: 'text',
  multiline: 'textarea', longtext: 'textarea', long_text: 'textarea',
  select: 'dropdown', combobox: 'dropdown',
  multiselect: 'multi_select', 'multi-select': 'multi_select',
  checkboxes: 'checklist', checkboxgroup: 'checklist',
  daterange: 'date_range', 'date-range': 'date_range',
  boolean: 'checkbox', toggle: 'checkbox',
  integer: 'number', decimal: 'number', float: 'number',
  upload: 'file', attachment: 'file', stars: 'rating',
  // people picker type names
  peoplepicker: 'peoplePicker', 'people-picker': 'peoplePicker', people_picker: 'peoplePicker',
  people: 'peoplePicker', person: 'peoplePicker', user: 'peoplePicker', users: 'peoplePicker',
  'person-or-group': 'peoplePicker', person_or_group: 'peoplePicker', personorgroup: 'peoplePicker',
  sharepointuser: 'peoplePicker', 'sharepoint-user': 'peoplePicker', 'sp.user': 'peoplePicker',
  principal: 'peoplePicker'
};

// Names (type, format or definition/$ref names) that mean "this is a person field"
const PICKER_NAMES = [
  'peoplepicker', 'people', 'person', 'user', 'users', 'sharepointuser', 'spuser',
  'sharepointperson', 'sharepointpeoplepicker', 'personorgroup', 'principal', 'userinfo'
];
const PLURAL_PICKER_NAMES = ['people', 'users'];

function isObj(v: any): boolean {
  return !!v && typeof v === 'object' && !Array.isArray(v);
}

function normalizeType(t: any): string {
  const s = String(t || 'text').toLowerCase();
  return TYPE_ALIASES[s] || s;
}

function uniqueId(ctx: IBuildCtx, base: string): string {
  let id = base;
  let n = 2;
  while (ctx.index[id]) { id = `${base}_${n}`; n++; }
  return id;
}

function register(ctx: IBuildCtx, node: IFormNode, parentId: string | undefined): void {
  ctx.index[node.id] = { node, parentId };
  if (node.kind === 'field') ctx.fields.push(node.field);
}

function toOptions(raw: any): IFieldOption[] {
  if (!Array.isArray(raw)) return [];
  return raw.map((o: any) => {
    if (isObj(o)) {
      const value = o.value !== undefined ? o.value : o.label;
      const label = o.label !== undefined ? o.label : (o.text !== undefined ? o.text : o.value);
      return { label: String(label), value: String(value) };
    }
    return { label: String(o), value: String(o) };
  });
}

// ---------------------------------------------------------------
// Styles that come from the schema
// ---------------------------------------------------------------

/** "background-color" -> "backgroundColor" (custom properties like --x are kept). */
function toCamel(key: string): string {
  return key.indexOf('--') === 0
    ? key
    : key.replace(/-([a-z])/g, (_m: string, c: string) => c.toUpperCase());
}

/**
 * Accepts an object ({ "backgroundColor": "#fff" }) or a CSS string
 * ("background-color: #fff; padding: 8px") and returns a style map.
 * Values that could load outside content (url(), expression(), javascript:) are dropped.
 */
function parseStyle(value: any): IStyleMap | undefined {
  const entries: Array<[string, any]> = [];

  if (typeof value === 'string') {
    value.split(';').forEach((part: string) => {
      const i = part.indexOf(':');
      if (i > 0) {
        entries.push([part.substring(0, i).trim(), part.substring(i + 1).trim()]);
      }
    });
  } else if (isObj(value)) {
    Object.keys(value).forEach((k) => entries.push([k, value[k]]));
  } else {
    return undefined;
  }

  const out: IStyleMap = {};
  let count = 0;

  entries.forEach((entry) => {
    const key = entry[0];
    const val = entry[1];

    if (!key) return;

    if (typeof val === 'number' && isFinite(val)) {
      out[toCamel(key)] = val;
      count++;
    } else if (typeof val === 'string' && val !== '') {
      if (/url\s*\(|expression\s*\(|javascript:/i.test(val)) return;
      out[toCamel(key)] = val;
      count++;
    }
  });

  return count > 0 ? out : undefined;
}

/** Reads "key", "x-key" (and "css" / "x-css" for "style") from a schema node. */
function readStyle(raw: any, key: string): IStyleMap | undefined {
  if (!isObj(raw)) return undefined;

  let v: any = raw[key];
  if (v === undefined) v = raw['x-' + key];
  if (v === undefined && key === 'style') {
    v = raw.css !== undefined ? raw.css : raw['x-css'];
  }

  return parseStyle(v);
}

function fieldStyles(raw: any): {
  style?: IStyleMap; labelStyle?: IStyleMap; controlStyle?: IStyleMap; helpStyle?: IStyleMap;
} {
  return {
    style: readStyle(raw, 'style'),
    labelStyle: readStyle(raw, 'labelStyle'),
    controlStyle: readStyle(raw, 'controlStyle'),
    helpStyle: readStyle(raw, 'helpStyle')
  };
}

// ---------------------------------------------------------------
// People Picker detection
// ---------------------------------------------------------------

function cleanName(v: any): string {
  return String(v === undefined || v === null ? '' : v).toLowerCase().replace(/[^a-z]/g, '');
}

function isPickerName(v: any): boolean {
  return PICKER_NAMES.indexOf(cleanName(v)) >= 0;
}

/** Last part of a "$ref": "#/definitions/User" -> "User". */
function refTail(raw: any): string {
  if (!isObj(raw) || typeof raw.$ref !== 'string') return '';
  const parts = raw.$ref.split('/');
  return parts[parts.length - 1];
}

/** All type-like names declared on a node (type may be a string or an array). */
function typeNames(raw: any): string[] {
  const t = isObj(raw) ? raw.type : undefined;
  return (Array.isArray(t) ? t : [t]).filter((x: any) => typeof x === 'string');
}

// An object that looks like a SharePoint person: { Id, Title, Email }
function hasPersonShape(s: any): boolean {
  if (!isObj(s) || !isObj(s.properties)) return false;
  const keys = Object.keys(s.properties).map((k) => k.toLowerCase());
  return keys.length <= 10 &&
    keys.indexOf('id') >= 0 &&
    keys.indexOf('title') >= 0 &&
    (keys.indexOf('email') >= 0 || keys.indexOf('loginname') >= 0);
}

function isPeoplePickerNode(raw: any, itemSchema?: any, refNames?: string[]): boolean {
  if (!isObj(raw)) return false;

  // explicit control
  const control = String(raw['x-control'] || raw.control || '').toLowerCase();
  if (control === 'peoplepicker' || control === 'people-picker' || control === 'people_picker') return true;

  // type name: "peoplePicker", "people", "person", "user", "users", ...
  if (typeNames(raw).some(isPickerName)) return true;
  if (itemSchema && typeNames(itemSchema).some(isPickerName)) return true;

  // format name
  if (isPickerName(raw.format) || (itemSchema && isPickerName(itemSchema.format))) return true;

  // $ref / definition name: "#/definitions/User"
  if ((refNames || []).some(isPickerName)) return true;

  // description mentions people picker
  if (typeof raw.description === 'string' && /people\s*-?\s*picker/i.test(raw.description)) return true;

  // a single person object, or an array of person objects
  return hasPersonShape(raw) || hasPersonShape(itemSchema);
}

function isPluralPicker(raw: any, itemSchema?: any): boolean {
  const names = typeNames(raw).concat(itemSchema ? typeNames(itemSchema) : []);
  return names.some((n) => PLURAL_PICKER_NAMES.indexOf(cleanName(n)) >= 0);
}

function getPeoplePickerConfig(raw: any): IPeoplePickerConfig {
  const config = raw['x-peoplePicker'] || raw.peoplePicker;
  if (!isObj(config)) return {};

  return {
    principalTypes: Array.isArray(config.principalTypes)
      ? config.principalTypes.map((x: any) => String(x))
      : undefined,
    allowMultiple: typeof config.allowMultiple === 'boolean' ? config.allowMultiple : undefined,
    personSelectionLimit: typeof config.personSelectionLimit === 'number'
      ? config.personSelectionLimit : undefined,
    showTooltip: typeof config.showTooltip === 'boolean' ? config.showTooltip : undefined,
    showProfilePicture: typeof config.showProfilePicture === 'boolean'
      ? config.showProfilePicture : undefined,
    profilePictureSize: typeof config.profilePictureSize === 'number'
      ? config.profilePictureSize : undefined,
    resolveDelay: typeof config.resolveDelay === 'number' ? config.resolveDelay : undefined,
    placeholder: typeof config.placeholder === 'string' ? config.placeholder : undefined,
    required: typeof config.required === 'boolean' ? config.required : undefined,
    ensureUser: typeof config.ensureUser === 'boolean' ? config.ensureUser : undefined
  };
}

// ---------------------------------------------------------------
// JSON Schema helpers ($ref, allOf)
// ---------------------------------------------------------------

function resolveRef(root: any, schema: any, depth: number): any {
  if (!isObj(schema) || typeof schema.$ref !== 'string' || schema.$ref.indexOf('#/') !== 0 || depth > 10) {
    return schema;
  }

  let cur: any = root;
  const parts = schema.$ref.substring(2).split('/').map(
    (p: string) => p.replace(/~1/g, '/').replace(/~0/g, '~')
  );
  for (let i = 0; i < parts.length; i++) {
    cur = isObj(cur) ? cur[parts[i]] : undefined;
  }
  if (!isObj(cur)) return schema;

  const siblings = Object.assign({}, schema);
  delete siblings.$ref;
  return resolveRef(root, Object.assign({}, cur, siblings), depth + 1);
}

function expandSchema(root: any, schema: any): any {
  const s = resolveRef(root, schema, 0);
  if (!isObj(s) || !Array.isArray(s.allOf)) return s;

  const merged: any = Object.assign({}, s);
  delete merged.allOf;
  merged.properties = Object.assign({}, merged.properties);
  merged.required = (merged.required || []).slice();

  s.allOf.forEach((part: any) => {
    const p = expandSchema(root, part);
    if (!isObj(p)) return;
    Object.assign(merged.properties, p.properties || {});
    merged.required = merged.required.concat(p.required || []);
    ['title', 'description', 'type'].forEach((k) => {
      if (merged[k] === undefined && p[k] !== undefined) merged[k] = p[k];
    });
  });
  return merged;
}

function parseSchemaProps(
  ctx: IBuildCtx, root: any, schemaIn: any, parentPath: string, parentId: string | undefined
): IFormNode[] {
  const schema = expandSchema(root, schemaIn);
  const props = isObj(schema) && isObj(schema.properties) ? schema.properties : {};
  const requiredList: string[] = Array.isArray(schema.required) ? schema.required : [];
  const out: IFormNode[] = [];

  Object.keys(props).forEach((key) => {
    const path = parentPath ? `${parentPath}.${key}` : key;
    const rawProp = props[key];
    const p = expandSchema(root, rawProp);
    if (!isObj(p)) return;

    // names of any $ref used by this property (checked before the ref is resolved away)
    const refNames = [refTail(rawProp), isObj(rawProp) ? refTail(rawProp.items) : ''];

    const id = uniqueId(ctx, path);
    const title = String(p.title || key);
    const description = typeof p.description === 'string' ? p.description : undefined;
    const required = requiredList.indexOf(key) >= 0;
    const itemSchema = isObj(p.items) ? expandSchema(root, p.items) : undefined;
    const typeList: any[] = Array.isArray(p.type) ? p.type.filter((t: any) => t !== 'null') : [p.type];
    const type: string = typeList[0] || (isObj(p.properties) ? 'object' : (p.items ? 'array' : 'string'));

    // IMPORTANT: People Picker must be detected BEFORE object/array handling,
    // otherwise a person object ({ Id, Title, Email }) becomes a group of text boxes.
    if (isPeoplePickerNode(p, itemSchema, refNames)) {
      const cfg = getPeoplePickerConfig(p);
      const desc = description || '';
      const isMulti = type === 'array' || cfg.allowMultiple === true || isPluralPicker(p, itemSchema) ||
        (/multi/i.test(desc) && !/single/i.test(desc));
      const limit = cfg.personSelectionLimit !== undefined
        ? cfg.personSelectionLimit
        : (isMulti ? (typeof p.maxItems === 'number' ? p.maxItems : 10) : 1);

      const pickerNode: IFormNode = {
        kind: 'field',
        id,
        field: {
          fieldId: id,
          label: title,
          type: 'peoplePicker',
          control: 'peoplePicker',
          peoplePicker: Object.assign({}, cfg, { allowMultiple: isMulti, personSelectionLimit: limit }),
          required: required || cfg.required === true,
          placeholder: cfg.placeholder,
          helpText: description,
          ...fieldStyles(p)
        }
      };
      register(ctx, pickerNode, parentId);
      out.push(pickerNode);
      return;
    }

    const isObjectType = type === 'object' || isObj(p.properties);
    const isObjectArray = type === 'array' && !!itemSchema &&
      (itemSchema.type === 'object' || isObj(itemSchema.properties));

    // ----- nested object / array of objects -> group -----
    if (isObjectType || isObjectArray) {
      const group: IFormGroup = {
        kind: 'group',
        id,
        title,
        description,
        children: [],
        style: readStyle(p, 'style'),
        legendStyle: readStyle(p, 'legendStyle')
      };
      register(ctx, group, parentId);
      group.children = parseSchemaProps(ctx, root, isObjectType ? p : itemSchema, path, id);
      out.push(group);
      return;
    }

    // ----- plain field -----
    let fieldType = 'text';
    let options: IFieldOption[] = [];

    if (Array.isArray(p.enum)) {
      fieldType = 'dropdown';
      options = toOptions(p.enum);
    } else if (type === 'array') {
      if (itemSchema && Array.isArray(itemSchema.enum)) {
        fieldType = 'checklist';
        options = toOptions(itemSchema.enum);
      } else {
        fieldType = 'list';
      }
    } else if (type === 'boolean') {
      fieldType = 'checkbox';
    } else if (type === 'integer' || type === 'number') {
      fieldType = 'number';
    } else if (p.format === 'date' || p.format === 'date-time') {
      fieldType = 'date';
    } else if (p.format === 'textarea' || (typeof p.maxLength === 'number' && p.maxLength > 200)) {
      fieldType = 'textarea';
    }

    const field: IFormField = {
      fieldId: id,
      label: title,
      type: fieldType,
      required,
      helpText: description,
      options,
      defaultValue: p.default,
      placeholder: Array.isArray(p.examples) && typeof p.examples[0] === 'string' ? p.examples[0] : undefined,
      validation: {
        min: typeof p.minimum === 'number' ? p.minimum : undefined,
        max: typeof p.maximum === 'number' ? p.maximum : undefined,
        minLength: typeof p.minLength === 'number' ? p.minLength : undefined,
        maxLength: typeof p.maxLength === 'number' ? p.maxLength : undefined,
        pattern: typeof p.pattern === 'string' ? p.pattern : undefined
      },
      ...fieldStyles(p)
    };
    const node: IFormNode = { kind: 'field', id, field };
    register(ctx, node, parentId);
    out.push(node);
  });

  return out;
}

// ---------------------------------------------------------------
// Custom format helpers (sections / fields / nested sections)
// ---------------------------------------------------------------

function hasContainer(raw: any): boolean {
  return CONTAINER_KEYS.some((k) => Array.isArray(raw[k]));
}

function buildField(raw: any, id: string, title: string, visibleWhen: IVisibleWhen | undefined): IFormField {
  const isPicker = isPeoplePickerNode(raw);
  let pickerCfg: IPeoplePickerConfig | undefined;

  if (isPicker) {
    const cfg = getPeoplePickerConfig(raw);
    const isMulti = cfg.allowMultiple === true || isPluralPicker(raw);
    pickerCfg = Object.assign({}, cfg, {
      allowMultiple: isMulti,
      personSelectionLimit: cfg.personSelectionLimit !== undefined
        ? cfg.personSelectionLimit
        : (isMulti ? 10 : 1)
    });
  }

  return {
    fieldId: id,
    label: title,
    type: isPicker ? 'peoplePicker' : normalizeType(raw.type),
    control: isPicker ? 'peoplePicker' : undefined,
    peoplePicker: pickerCfg,
    required: raw.required === true || (pickerCfg !== undefined && pickerCfg.required === true),
    placeholder: typeof raw.placeholder === 'string'
      ? raw.placeholder
      : (pickerCfg ? pickerCfg.placeholder : undefined),
    helpText: typeof raw.helpText === 'string'
      ? raw.helpText
      : (typeof raw.description === 'string' ? raw.description : undefined),
    options: toOptions(raw.options || raw.choices || raw.values),
    validation: isObj(raw.validation) ? Object.assign({}, raw.validation) : undefined,
    visibleWhen,
    defaultValue: raw.defaultValue !== undefined ? raw.defaultValue : raw.default,
    min: typeof raw.min === 'number' ? raw.min : undefined,
    max: typeof raw.max === 'number' ? raw.max : undefined,
    ...fieldStyles(raw)
  };
}

function parseCustomNode(
  ctx: IBuildCtx, raw: any, position: number, parentId: string | undefined
): IFormNode | undefined {
  if (!isObj(raw)) return undefined;

  const base = String(
    raw.fieldId || raw.sectionId || raw.groupId || raw.id || raw.key || raw.name || `${parentId || 'item'}_${position}`
  );
  const id = uniqueId(ctx, base);
  const title = String(raw.label || raw.sectionTitle || raw.groupTitle || raw.title || raw.name || base);
  const visibleWhen = isObj(raw.visibleWhen) ? raw.visibleWhen as IVisibleWhen : undefined;
  const isGroup = !isPeoplePickerNode(raw) &&
    (hasContainer(raw) || GROUP_TYPES.indexOf(String(raw.type || '').toLowerCase()) >= 0);

  if (isGroup) {
    const group: IFormGroup = {
      kind: 'group',
      id,
      title,
      description: typeof raw.description === 'string' ? raw.description : undefined,
      visibleWhen,
      children: [],
      style: readStyle(raw, 'style'),
      legendStyle: readStyle(raw, 'legendStyle')
    };
    register(ctx, group, parentId);
    group.children = parseChildren(ctx, raw, id);
    return group;
  }

  const node: IFormNode = { kind: 'field', id, field: buildField(raw, id, title, visibleWhen) };
  register(ctx, node, parentId);
  return node;
}

function parseChildren(ctx: IBuildCtx, raw: any, parentId: string | undefined): IFormNode[] {
  const out: IFormNode[] = [];
  CONTAINER_KEYS.forEach((key) => {
    const arr = raw[key];
    if (!Array.isArray(arr)) return;
    arr.forEach((item: any, i: number) => {
      const node = parseCustomNode(ctx, item, i, parentId);
      if (node) out.push(node);
    });
  });
  return out;
}

function looksLikeField(o: any): boolean {
  return isObj(o) && (
    typeof o.fieldId === 'string' ||
    (typeof o.type === 'string' && (typeof o.label === 'string' || typeof o.name === 'string'))
  );
}

// Last resort: look anywhere in the file for a schema or a list of field-like objects.
function deepFind(ctx: IBuildCtx, rawRoot: any, node: any, depth: number): IFormNode[] {
  if (!isObj(node) || depth > 5) return [];

  if (isObj(node.properties) && Object.keys(node.properties).length > 0) {
    return parseSchemaProps(ctx, rawRoot, node, '', undefined);
  }

  const keys = Object.keys(node);
  for (let i = 0; i < keys.length; i++) {
    const v = node[keys[i]];

    if (Array.isArray(v) && v.some((x: any) => looksLikeField(x) || (isObj(x) && hasContainer(x)))) {
      const out: IFormNode[] = [];
      v.forEach((item: any, idx: number) => {
        const n = parseCustomNode(ctx, item, idx, undefined);
        if (n) out.push(n);
      });
      if (ctx.fields.length > 0) return out;
    }

    if (isObj(v)) {
      const found = deepFind(ctx, rawRoot, v, depth + 1);
      if (ctx.fields.length > 0) return found;
    }
  }
  return [];
}

// ---------------------------------------------------------------
// Public API
// ---------------------------------------------------------------

export function parseDefinition(raw: any): IFormDefinition {
  let ctx: IBuildCtx = { fields: [], index: {} };
  const root = expandSchema(raw, raw);
  let nodes: IFormNode[];

  if (isObj(root) && isObj(root.properties) && Object.keys(root.properties).length > 0) {
    nodes = parseSchemaProps(ctx, raw, root, '', undefined);
  } else {
    nodes = parseChildren(ctx, isObj(root) ? root : {}, undefined);
  }

  if (ctx.fields.length === 0) {
    ctx = { fields: [], index: {} };
    nodes = deepFind(ctx, raw, raw, 0);
  }

  const actions: IFormAction[] = Array.isArray(raw.actions)
    ? raw.actions.filter((a: any) => isObj(a)).map((a: any) => ({
        actionId: String(a.actionId || a.id || a.type || a.label),
        label: String(a.label || a.type || 'Submit'),
        type: String(a.type || 'submit').toLowerCase(),
        style: readStyle(a, 'style')
      }))
    : [];

  // Named style parts: "styles": { "label": { ... }, "primaryButton": { ... } }
  const rawStyles: any = isObj(raw.styles) ? raw.styles : (isObj(raw['x-styles']) ? raw['x-styles'] : {});
  const styles: IFormStyles = {};

  STYLE_PARTS.forEach((part) => {
    const s = parseStyle(rawStyles[part]);
    if (s) styles[part as keyof IFormStyles] = s;
  });

  return {
    formId: typeof raw.formId === 'string' ? raw.formId : undefined,
    formName: String(raw.formName || raw.title || raw.name || 'Form'),
    version: typeof raw.version === 'string' ? raw.version : undefined,
    nodes,
    fields: ctx.fields,
    index: ctx.index,
    actions,
    style: readStyle(raw, 'style'),
    styles
  };
}

export function getInitialValues(def: IFormDefinition): IFieldValues {
  const values: IFieldValues = {};
  def.fields.forEach((f) => {
    if (f.type === 'peoplePicker') return; // pickers start empty
    if (f.defaultValue !== undefined && f.defaultValue !== null) values[f.fieldId] = f.defaultValue;
  });
  return values;
}

function same(a: unknown, b: unknown): boolean {
  return a === b || (a !== undefined && b !== undefined && a !== null && b !== null && String(a) === String(b));
}

function inList(list: unknown[], value: unknown): boolean {
  return list.some((x) => same(x, value));
}

/** A node is visible when it, its parent groups and its controlling field are all visible. */
export function isNodeVisible(def: IFormDefinition, id: string, values: IFieldValues, depth = 0): boolean {
  if (depth > 20) return false;

  const entry = def.index[id];
  if (!entry) return false;

  if (entry.parentId && !isNodeVisible(def, entry.parentId, values, depth + 1)) return false;

  const cond = entry.node.kind === 'group' ? entry.node.visibleWhen : entry.node.field.visibleWhen;
  if (!cond) return true;

  if (!def.index[cond.field]) return false;
  if (!isNodeVisible(def, cond.field, values, depth + 1)) return false;

  const actual = values[cond.field];
  const op = cond.operator || 'equals';

  if (op === 'notEquals') {
    return Array.isArray(actual) ? !inList(actual, cond.value) : !same(actual, cond.value);
  }
  if (op === 'contains') {
    return Array.isArray(actual)
      ? inList(actual, cond.value)
      : String(actual || '').indexOf(String(cond.value)) >= 0;
  }
  if (op === 'in') {
    return Array.isArray(cond.value) ? inList(cond.value as unknown[], actual) : false;
  }
  return Array.isArray(actual) ? inList(actual, cond.value) : same(actual, cond.value);
}

function isEmpty(field: IFormField, v: unknown): boolean {
  if (v === undefined || v === null) return true;
  if (typeof v === 'string') return v.trim() === '';
  if (Array.isArray(v)) return v.length === 0;
  if (field.type === 'checkbox') return v !== true;
  if (field.type === 'date_range') {
    const r = v as IDateRange;
    return !r.start || !r.end;
  }
  return false;
}

/** Returns the values of visible fields only (empty values are dropped). */
export function collectValues(def: IFormDefinition, values: IFieldValues): IFieldValues {
  const out: IFieldValues = {};

  def.fields.forEach((f) => {
    if (!isNodeVisible(def, f.fieldId, values)) return;

    const v = values[f.fieldId];
    if (v === undefined || v === null) return;
    if (typeof v === 'string' && v.trim() === '') return;
    if (Array.isArray(v) && v.length === 0) return;

    if (f.type === 'date_range') {
      const r = v as IDateRange;
      if (!r.start && !r.end) return;
    }

    // People picker: save in the { Id, Title, Email } shape
    if (f.type === 'peoplePicker' && Array.isArray(v)) {
      const people = (v as any[]).map((p) => ({ Id: p.id, Title: p.text, Email: p.email }));
      const multi = !!(f.peoplePicker && f.peoplePicker.allowMultiple);
      out[f.fieldId] = multi ? people : people[0];
      return;
    }

    if (f.type === 'list' && typeof v === 'string') {
      out[f.fieldId] = v.split(',').map((x) => x.trim()).filter((x) => x !== '');
      return;
    }

    out[f.fieldId] = v;
  });

  return out;
}

/** Returns a list of error messages (empty when the form is valid). */
export function validateForm(def: IFormDefinition, values: IFieldValues): string[] {
  const errors: string[] = [];

  def.fields.forEach((f) => {
    if (!isNodeVisible(def, f.fieldId, values)) return;

    const v = values[f.fieldId];
    const label = f.label || f.fieldId;
    const rules = f.validation || {};

    if (f.required && isEmpty(f, v)) {
      errors.push(`${label} is required.`);
      return;
    }
    if (isEmpty(f, v)) return;

    if (f.type === 'number') {
      const n = Number(v);
      const min = rules.min !== undefined ? rules.min : f.min;
      const max = rules.max !== undefined ? rules.max : f.max;
      if (!isFinite(n)) errors.push(`${label} must be a number.`);
      if (min !== undefined && n < min) errors.push(`${label} must be at least ${min}.`);
      if (max !== undefined && n > max) errors.push(`${label} must be at most ${max}.`);
    }

    if (typeof v === 'string') {
      if (rules.minLength !== undefined && v.length < rules.minLength) {
        errors.push(`${label} must be at least ${rules.minLength} characters.`);
      }
      if (rules.maxLength !== undefined && v.length > rules.maxLength) {
        errors.push(`${label} must be at most ${rules.maxLength} characters.`);
      }
      if (rules.pattern) {
        try {
          // The pattern comes from the admin-controlled schema file, so it must be built at runtime.
          // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
          const regex = new RegExp(rules.pattern);
          if (!regex.test(v)) errors.push(`${label} is not in the expected format.`);
        } catch {
          // invalid pattern in the schema: ignore it
        }
      }
    }

    if (f.type === 'peoplePicker' && Array.isArray(v)) {
      const limit = f.peoplePicker && f.peoplePicker.personSelectionLimit;
      if (limit !== undefined && v.length > limit) {
        errors.push(`${label}: you can select at most ${limit} ${limit === 1 ? 'person' : 'people'}.`);
      }
    }

    if (f.type === 'date_range') {
      const r = v as IDateRange;
      if (r.start && r.end && r.end < r.start) {
        errors.push(`${label}: end date must be on or after the start date.`);
      }
    }

    if (f.type === 'file' && Array.isArray(v)) {
      const files = v as IFileInfo[];
      if (rules.maxFiles !== undefined && files.length > rules.maxFiles) {
        errors.push(`${label}: you can attach at most ${rules.maxFiles} file(s).`);
      }
      files.forEach((file) => {
        const lower = file.name.toLowerCase();
        if (rules.allowedExtensions && rules.allowedExtensions.length > 0) {
          const ok = rules.allowedExtensions.some(
            (ext) => lower.substring(lower.length - ext.length) === ext.toLowerCase()
          );
          if (!ok) {
            errors.push(`${label}: ${file.name} is not an allowed type (${rules.allowedExtensions.join(', ')}).`);
          }
        }
        if (rules.maxFileSizeMB !== undefined && file.size > rules.maxFileSizeMB * 1024 * 1024) {
          errors.push(`${label}: ${file.name} is larger than ${rules.maxFileSizeMB} MB.`);
        }
      });
    }
  });

  return errors;
}