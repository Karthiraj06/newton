/* Reconstructed from usage in Formfull.ts and SchemaDrivenForm.tsx, with schema-driven styling */

export interface IFieldOption {
  label: string;
  value: string;
}

export interface IFileInfo {
  name: string;
  size: number;
  type?: string;
  /** The real File object, kept so it can be uploaded as an attachment. */
  file?: File;
}

export interface IDateRange {
  start?: string;
  end?: string;
}

export interface IFieldValues {
  [fieldId: string]: unknown;
}

/** CSS properties read from the schema, e.g. { backgroundColor: '#fff', padding: 8 } */
export interface IStyleMap {
  [property: string]: string | number;
}

/** Named parts of the form that the schema can style (schema key: "styles"). */
export interface IFormStyles {
  container?: IStyleMap;
  title?: IStyleMap;
  group?: IStyleMap;
  legend?: IStyleMap;
  field?: IStyleMap;
  label?: IStyleMap;
  help?: IStyleMap;
  control?: IStyleMap;
  option?: IStyleMap;
  dateRange?: IStyleMap;
  rating?: IStyleMap;
  ratingButton?: IStyleMap;
  ratingSelected?: IStyleMap;
  actions?: IStyleMap;
  primaryButton?: IStyleMap;
  secondaryButton?: IStyleMap;
  error?: IStyleMap;
  success?: IStyleMap;
}

export interface IPeoplePickerConfig {
  principalTypes?: Array<string | number>;
  allowMultiple?: boolean;
  personSelectionLimit?: number;
  showTooltip?: boolean;
  showProfilePicture?: boolean;
  profilePictureSize?: number;
  resolveDelay?: number;
  placeholder?: string;
  required?: boolean;
  ensureUser?: boolean;
}

export interface IVisibleWhen {
  field: string;
  operator?: 'equals' | 'notEquals' | 'contains' | 'in';
  value?: unknown;
}

export interface IFieldValidation {
  min?: number;
  max?: number;
  minLength?: number;
  maxLength?: number;
  pattern?: string;
  maxFiles?: number;
  allowedExtensions?: string[];
  maxFileSizeMB?: number;
}

export interface IFormField {
  fieldId: string;
  label: string;
  type: string;
  control?: string;
  peoplePicker?: IPeoplePickerConfig;
  required?: boolean;
  placeholder?: string;
  helpText?: string;
  options?: IFieldOption[];
  validation?: IFieldValidation;
  visibleWhen?: IVisibleWhen;
  defaultValue?: unknown;
  min?: number;
  max?: number;

  // styling from the schema
  style?: IStyleMap;
  labelStyle?: IStyleMap;
  controlStyle?: IStyleMap;
  helpStyle?: IStyleMap;
}

export interface IFormGroup {
  kind: 'group';
  id: string;
  title: string;
  description?: string;
  visibleWhen?: IVisibleWhen;
  children: IFormNode[];

  // styling from the schema
  style?: IStyleMap;
  legendStyle?: IStyleMap;
}

export interface IFormFieldNode {
  kind: 'field';
  id: string;
  field: IFormField;
}

export type IFormNode = IFormGroup | IFormFieldNode;

export interface IIndexEntry {
  node: IFormNode;
  parentId?: string;
}

export interface IFormAction {
  actionId: string;
  label: string;
  /** 'submit' | 'save' | ... */
  type: string;
  style?: IStyleMap;
}

export interface IFormDefinition {
  formId?: string;
  formName: string;
  version?: string;
  nodes: IFormNode[];
  fields: IFormField[];
  index: { [id: string]: IIndexEntry };
  actions: IFormAction[];

  // styling from the schema
  style?: IStyleMap;
  styles?: IFormStyles;
}