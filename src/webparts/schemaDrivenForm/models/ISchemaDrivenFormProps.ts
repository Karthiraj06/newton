import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISchemaDrivenFormProps {
  context: WebPartContext;
  libraryTitle: string;
  schemaFileServerRelativeUrl: string;
  submissionLibraryId: string;
}

export interface IPeoplePickerConfig {
  /** 1 = user, 2 = distribution list, 4 = security group, 8 = SharePoint group */
  principalTypes?: number[];
  allowMultiple?: boolean;
  personSelectionLimit?: number;
  showTooltip?: boolean;
  resolveDelay?: number;
  placeholder?: string;
  required?: boolean;
  /** Resolve the person to a SharePoint user ID when selected. Default: true. */
  ensureUser?: boolean;
  /** Show each person's profile picture. Default: false. */
  showProfilePicture?: boolean;
  /** Profile picture size in pixels. Default: 28. */
  profilePictureSize?: number;
}

export interface IJsonSchema {
  $schema?: string;
  $id?: string;
  $ref?: string;
  title?: string;
  type?: string | string[];
  properties?: { [name: string]: IJsonSchema };
  required?: string[];
  items?: IJsonSchema | IJsonSchema[];
  enum?: unknown[];
  const?: unknown;
  format?: string;
  default?: unknown;
  description?: string;
  minimum?: number;
  maximum?: number;
  minLength?: number;
  maxLength?: number;
  pattern?: string;
  minItems?: number;
  maxItems?: number;
  additionalProperties?: boolean | IJsonSchema;

  // Custom rendering configuration
  'x-control'?: string;
  'x-peoplePicker'?: IPeoplePickerConfig;

  [keyword: string]: unknown;
}

export interface ISchemaPropertyInfo {
  path: string;
  title: string;
  type: string;
  required: boolean;
  format?: string;
  description?: string;

  // Custom control identification
  control?: string;

  // People Picker configuration
  peoplePicker?: IPeoplePickerConfig;
}