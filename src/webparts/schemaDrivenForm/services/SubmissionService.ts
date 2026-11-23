
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  ISPHttpClientOptions,
  MSGraphClientV3
} from '@microsoft/sp-http';

export interface ISubmissionResult {
  reference: string;
  itemId: number;
  itemUrl?: string;
  attachmentsUploaded: number;
  failedAttachments: string[];
  emailSent: boolean;
  emailError?: string;
}

interface IListInfo {
  Title: string;
  EnableAttachments?: boolean;
  RootFolder?: {
    ServerRelativeUrl: string;
  };
}

interface ISharePointField {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  ReadOnlyField?: boolean;
  Hidden?: boolean;
  LookupList?: string;
  LookupField?: string;
}

interface IListItemResponse {
  Id: number;
}

interface IListItemCreateResponse {
  d?: IListItemResponse;
  Id?: number;
}

interface IFieldValue {
  name: string;
  value: unknown;
}

interface IPendingFile {
  name: string;
  file: File;
}

export class SubmissionService {
  private readonly context: WebPartContext;
  private readonly userIdCache: { [key: string]: number } = {};

  // Testing recipient only
  private readonly testEmailRecipient =
    'karthiraj@infinitiresearch.com';

  public constructor(context: WebPartContext) {
    this.context = context;
  }

  private get webUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  private getHeaders(contentType?: string): { [key: string]: string } {
    return {
      Accept: 'application/json;odata=nometadata',
      ...(contentType ? { 'Content-Type': contentType } : {})
    };
  }

  private escapeXml(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  private escapeOData(value: string): string {
    return value.replace(/'/g, "''");
  }

  private escapeHtml(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // ==================================================
  // EMAIL HELPERS
  // ==================================================

  private formatEmailValue(value: unknown): string {
    if (value === null || value === undefined) {
      return '';
    }

    if (Array.isArray(value)) {
      return value
        .map(item => this.formatEmailValue(item))
        .filter(item => item !== '')
        .join(', ');
    }

    if (value instanceof Date) {
      return value.toLocaleString();
    }

    if (typeof value === 'object') {
      const obj = value as { [key: string]: unknown };

      const displayName =
        obj.displayName ||
        obj.DisplayName ||
        obj.name ||
        obj.title ||
        obj.text;

      const email =
        obj.email ||
        obj.Email ||
        obj.EMail ||
        obj.secondaryText;

      if (displayName && email) {
        return `${String(displayName)} (${String(email)})`;
      }

      if (displayName) {
        return String(displayName);
      }

      if (email) {
        return String(email);
      }

      try {
        return JSON.stringify(value);
      } catch {
        return '[Object value]';
      }
    }

    return String(value);
  }

  private getEmailErrorMessage(error: unknown): string {
    if (error instanceof Error) {
      const graphError = error as Error & {
        statusCode?: number;
        code?: string;
        body?: string;
      };

      return [
        graphError.message,
        graphError.statusCode
          ? `Status: ${graphError.statusCode}`
          : '',
        graphError.code || '',
        graphError.body || ''
      ]
        .filter(value => value !== '')
        .join(' | ');
    }

    if (error && typeof error === 'object') {
      try {
        return JSON.stringify(error);
      } catch {
        return 'Unknown Microsoft Graph error';
      }
    }

    return String(error);
  }

  /**
   * Sends submission details using Microsoft Graph.
   */
 
private async sendSubmissionEmail(
  values: { [key: string]: unknown },
  reference: string,
  itemId: number,
  itemUrl: string,
  formTitle: string
): Promise<void> {

  console.log('Getting Microsoft Graph client...');

  const client: MSGraphClientV3 =
    await this.context.msGraphClientFactory.getClient('3');

  console.log('Microsoft Graph client initialized.');

  let rows = '';

  Object.keys(values).forEach((key) => {
    const formattedValue = this.formatEmailValue(values[key]);

    rows += `
      <tr>
        <td style="padding:10px;border:1px solid #ddd;font-weight:bold;">
          ${this.escapeHtml(key)}
        </td>
        <td style="padding:10px;border:1px solid #ddd;">
          ${this.escapeHtml(formattedValue || '-')}
        </td>
      </tr>
    `;
  });

  const submittedBy =
    this.context.pageContext.user.displayName || '';

  const emailBody = `
    <div style="font-family:Arial,sans-serif;color:#323130;">
      <h2>New Form Submission</h2>
      <p>A new form submission has been received.</p>

      <p><strong>Form:</strong> ${this.escapeHtml(formTitle)}</p>
      <p><strong>Reference:</strong> ${this.escapeHtml(reference)}</p>
      <p><strong>Item ID:</strong> ${itemId}</p>
      <p><strong>Submitted By:</strong> ${this.escapeHtml(submittedBy)}</p>

      <h3>Submitted Form Values</h3>

      <table style="border-collapse:collapse;width:100%;">
        <thead>
          <tr>
            <th style="padding:10px;border:1px solid #ddd;text-align:left;">Field</th>
            <th style="padding:10px;border:1px solid #ddd;text-align:left;">Submitted Value</th>
          </tr>
        </thead>
        <tbody>
          ${rows}
        </tbody>
      </table>

      <p style="margin-top:20px;">
        <a href="${this.escapeHtml(itemUrl)}">View SharePoint Submission</a>
      </p>
    </div>
  `;

  console.log('Sending email to:', this.testEmailRecipient);

  try {
    const response = await client.api('/me/sendMail').post({
      message: {
        subject: `New Form Submission - ${reference}`,
        body: {
          contentType: 'HTML',
          content: emailBody
        },
        toRecipients: [
          {
            emailAddress: {
              address: this.testEmailRecipient
            }
          }
        ]
      },
      saveToSentItems: true
    });

    console.log('Graph sendMail response:', response);
    console.log('Email request accepted by Microsoft Graph.');

  } catch (error: unknown) {
    console.error('Microsoft Graph sendMail failed:', error);

    const graphError = error as {
      message?: string;
      statusCode?: number;
      code?: string;
      body?: unknown;
    };

    console.error('Graph status code:', graphError?.statusCode);
    console.error('Graph error code:', graphError?.code);
    console.error('Graph error message:', graphError?.message);
    console.error('Graph error body:', graphError?.body);

    throw error;
  }
}


  // ==================================================
  // SHAREPOINT FIELD HELPERS
  // ==================================================

  private getInternalName(fieldName: string): string {
    let name = fieldName
      .trim()
      .replace(/[^a-zA-Z0-9_]/g, '_')
      .replace(/_+/g, '_');

    if (!name) {
      name = 'CustomField';
    }

    if (!/^[a-zA-Z_]/.test(name)) {
      name = `Field_${name}`;
    }

    return name.substring(0, 32);
  }

  private findField(
    fields: ISharePointField[],
    fieldName: string
  ): ISharePointField | undefined {

    const lower = fieldName.trim().toLowerCase();
    const internalLower =
      this.getInternalName(fieldName).toLowerCase();

    return (
      fields.find(f => f.InternalName.toLowerCase() === lower) ||
      fields.find(f => f.InternalName.toLowerCase() === internalLower) ||
      fields.find(f => f.Title.toLowerCase() === lower)
    );
  }

  private getFieldType(
    fieldName: string,
    value: unknown
  ): string {

    if (fieldName === 'SubmittedOn') {
      return 'DateTime';
    }

    if (typeof value === 'boolean') {
      return 'Boolean';
    }

    if (typeof value === 'number') {
      return 'Number';
    }

    if (value !== null && typeof value === 'object') {
      return 'Note';
    }

    if (typeof value === 'string' && value.length > 255) {
      return 'Note';
    }

    return 'Text';
  }

  // ==================================================
  // VALUE HELPERS
  // ==================================================

  private isEmpty(value: unknown): boolean {
    return (
      value === null ||
      value === undefined ||
      (typeof value === 'string' && value.trim() === '') ||
      (Array.isArray(value) && value.length === 0)
    );
  }

  private toArray(value: unknown): unknown[] {
    return Array.isArray(value) ? value : [value];
  }

  private toText(value: unknown): string {
    if (value === null || value === undefined) {
      return '';
    }

    if (
      typeof value === 'string' ||
      typeof value === 'number' ||
      typeof value === 'boolean'
    ) {
      return String(value);
    }

    if (typeof value === 'object') {
      const obj = value as { [key: string]: unknown };

      const candidates = [
        'text', 'label', 'name', 'title',
        'value', 'key', 'displayName', 'lookupValue'
      ];

      for (const key of candidates) {
        const v = obj[key];

        if (typeof v === 'string' && v) {
          return v;
        }

        if (typeof v === 'number') {
          return String(v);
        }
      }

      return JSON.stringify(value);
    }

    return String(value);
  }

  private toNumericId(value: unknown): number | undefined {
    if (
      typeof value === 'number' &&
      isFinite(value) &&
      value > 0
    ) {
      return value;
    }

    if (
      typeof value === 'string' &&
      /^\d+$/.test(value.trim())
    ) {
      return parseInt(value.trim(), 10);
    }

    return undefined;
  }

  // ==================================================
  // FILE HELPERS
  // ==================================================

  private isFileEntry(
    value: unknown
  ): value is { name: string; file: File } {

    if (!value || typeof value !== 'object') {
      return false;
    }

    const o = value as {
      name?: unknown;
      file?: unknown;
    };

    return (
      typeof o.name === 'string' &&
      typeof File !== 'undefined' &&
      o.file instanceof File
    );
  }

  private extractFiles(
    values: { [key: string]: unknown }
  ): {
    clean: { [key: string]: unknown };
    files: IPendingFile[];
  } {

    const clean: { [key: string]: unknown } = {};
    const files: IPendingFile[] = [];

    Object.keys(values).forEach(key => {
      const v = values[key];

      if (
        Array.isArray(v) &&
        v.length > 0 &&
        v.some(x => this.isFileEntry(x))
      ) {
        const names: string[] = [];

        v.forEach(x => {
          if (this.isFileEntry(x)) {
            files.push({
              name: x.name,
              file: x.file
            });

            names.push(x.name);
          }
        });

        clean[key] = names.join(', ');
      } else {
        clean[key] = v;
      }
    });

    return { clean, files };
  }

  private uniqueName(
    name: string,
    used: { [name: string]: boolean }
  ): string {

    let candidate = name;
    let n = 2;

    const dot = name.lastIndexOf('.');
    const stem = dot > 0
      ? name.substring(0, dot)
      : name;

    const ext = dot > 0
      ? name.substring(dot)
      : '';

    while (used[candidate.toLowerCase()]) {
      candidate = `${stem}_${n}${ext}`;
      n++;
    }

    used[candidate.toLowerCase()] = true;

    return candidate;
  }

  private async uploadAttachment(
    listId: string,
    itemId: number,
    name: string,
    file: File
  ): Promise<void> {

    const safeName =
      encodeURIComponent(this.escapeOData(name));

    const buffer = await file.arrayBuffer();

    const response = await this.context.spHttpClient.post(
      `${this.webUrl}/_api/web/lists(guid'${this.escapeOData(listId)}')` +
      `/items(${itemId})/AttachmentFiles/add(FileName='${safeName}')`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        },
        body: buffer
      }
    );

    if (!response.ok) {
      throw new Error(
        `${response.status}: ${await response.text()}`
      );
    }
  }

  // ==================================================
  // PEOPLE PICKER
  // ==================================================

  private async ensureUserId(
    loginOrEmail: string
  ): Promise<number | undefined> {

    let logonName = loginOrEmail.trim();

    if (!logonName) {
      return undefined;
    }

    if (
      logonName.indexOf('@') > -1 &&
      logonName.indexOf('|') === -1
    ) {
      logonName =
        `i:0#.f|membership|${logonName}`;
    }

    const cacheKey = logonName.toLowerCase();

    if (this.userIdCache[cacheKey]) {
      return this.userIdCache[cacheKey];
    }

    const response = await this.context.spHttpClient.post(
      `${this.webUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({ logonName })
      }
    );

    if (!response.ok) {
      console.warn(
        `Could not resolve user "${loginOrEmail}" (${response.status}): ` +
        await response.text()
      );

      return undefined;
    }

    const user = await response.json() as {
      Id?: number;
    };

    if (user.Id) {
      this.userIdCache[cacheKey] = user.Id;
    }

    return user.Id;
  }

  private async resolveUserIds(
    value: unknown
  ): Promise<number[]> {

    const ids: number[] = [];

    for (const item of this.toArray(value)) {
      if (this.isEmpty(item)) {
        continue;
      }

      const numeric = this.toNumericId(item);

      if (numeric) {
        ids.push(numeric);
        continue;
      }

      const candidates: string[] = [];

      if (typeof item === 'string') {
        candidates.push(item);
      } else if (
        typeof item === 'object' &&
        item !== null
      ) {
        const obj = item as {
          [key: string]: unknown;
        };

        const objId =
          this.toNumericId(obj.id) ||
          this.toNumericId(obj.Id) ||
          this.toNumericId(obj.ID);

        if (objId) {
          ids.push(objId);
          continue;
        }

        [
          obj.loginName,
          obj.LoginName,
          obj.key,
          obj.email,
          obj.Email,
          obj.EMail,
          obj.secondaryText,
          obj.upn
        ].forEach(c => {
          if (typeof c === 'string' && c) {
            candidates.push(c);
          }
        });
      }

      for (const candidate of candidates) {
        const id = await this.ensureUserId(candidate);

        if (id) {
          ids.push(id);
          break;
        }
      }
    }

    return ids;
  }

  // ==================================================
  // LOOKUP HELPERS
  // ==================================================

  private async findLookupItemId(
    field: ISharePointField,
    text: string
  ): Promise<number | undefined> {

    if (!field.LookupList) {
      return undefined;
    }

    const listGuid =
      field.LookupList.replace(/[{}]/g, '');

    const lookupField =
      field.LookupField || 'Title';

    const url =
      `${this.webUrl}/_api/web/lists(guid'${listGuid}')/items` +
      `?$select=Id&$filter=${lookupField} eq '${this.escapeOData(text)}'&$top=1`;

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: this.getHeaders()
      }
    );

    if (!response.ok) {
      console.warn(
        `Could not resolve lookup value "${text}" (${response.status}).`
      );

      return undefined;
    }

    const data = await response.json() as {
      value?: Array<{ Id: number }>;
    };

    return data.value && data.value.length > 0
      ? data.value[0].Id
      : undefined;
  }

  private async resolveLookupIds(
    field: ISharePointField,
    value: unknown
  ): Promise<number[]> {

    const ids: number[] = [];

    for (const item of this.toArray(value)) {
      if (this.isEmpty(item)) {
        continue;
      }

      let id = this.toNumericId(item);

      if (
        !id &&
        typeof item === 'object' &&
        item !== null
      ) {
        const obj = item as {
          [key: string]: unknown;
        };

        id =
          this.toNumericId(obj.id) ||
          this.toNumericId(obj.Id) ||
          this.toNumericId(obj.ID) ||
          this.toNumericId(obj.key) ||
          this.toNumericId(obj.lookupId) ||
          this.toNumericId(obj.value);
      }

      if (!id) {
        const text = this.toText(item);

        if (text) {
          id = await this.findLookupItemId(field, text);
        }
      }

      if (id) {
        ids.push(id);
      }
    }

    return ids;
  }

  

  // ==================================================
  // CONVERT FORM VALUE TO SHAREPOINT VALUE
  // ==================================================

  private async buildFieldValue(
    field: ISharePointField,
    value: unknown
  ): Promise<IFieldValue | undefined> {

    if (this.isEmpty(value)) {
      return undefined;
    }

    const name = field.InternalName;

    switch (field.TypeAsString) {
      case 'User':
      case 'UserMulti': {
        const ids = await this.resolveUserIds(value);

        if (ids.length === 0) {
          return undefined;
        }

        return {
          name: `${name}Id`,
          value: field.TypeAsString === 'UserMulti'
            ? ids
            : ids[0]
        };
      }

      case 'Lookup':
      case 'LookupMulti': {
        const ids = await this.resolveLookupIds(
          field,
          value
        );

        if (ids.length === 0) {
          return undefined;
        }

        return {
          name: `${name}Id`,
          value: field.TypeAsString === 'LookupMulti'
            ? ids
            : ids[0]
        };
      }

      case 'MultiChoice': {
        const choices = this.toArray(value)
          .map(v => this.toText(v))
          .filter(v => v !== '');

        return choices.length > 0
          ? { name, value: choices }
          : undefined;
      }

      case 'Choice':
        return {
          name,
          value: this.toText(this.toArray(value)[0])
        };

      case 'DateTime': {
        const date = value instanceof Date
          ? value
          : new Date(String(value));

        return isNaN(date.getTime())
          ? undefined
          : {
              name,
              value: date.toISOString()
            };
      }

      case 'Boolean': {
        if (typeof value === 'boolean') {
          return { name, value };
        }

        const text = String(value).toLowerCase();

        return {
          name,
          value:
            text === 'true' ||
            text === 'yes' ||
            text === '1'
        };
      }

      case 'Number':
      case 'Currency':
      case 'Integer': {
        const num = Number(value);

        return isFinite(num)
          ? { name, value: num }
          : undefined;
      }

      case 'URL': {
        if (typeof value === 'string') {
          return {
            name,
            value: {
              Url: value,
              Description: value
            }
          };
        }

        const obj = value as {
          [key: string]: unknown;
        };

        const url = this.toText(
          obj.Url || obj.url || obj.href || ''
        );

        return url
          ? {
              name,
              value: {
                Url: url,
                Description: this.toText(
                  obj.Description ||
                  obj.description ||
                  url
                )
              }
            }
          : undefined;
      }

      case 'Note':
      case 'Text':
        return {
          name,
          value: typeof value === 'object'
            ? JSON.stringify(value)
            : String(value)
        };

      default:
        return {
          name,
          value: typeof value === 'object'
            ? JSON.stringify(value)
            : value
        };
    }
  }

  // ==================================================
  // SHAREPOINT API
  // ==================================================

  private async getListInfo(
    listId: string
  ): Promise<IListInfo> {

    const url =
      `${this.webUrl}/_api/web/lists(guid'${this.escapeOData(listId)}')` +
      '?$select=Title,EnableAttachments,RootFolder/ServerRelativeUrl' +
      '&$expand=RootFolder';

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: this.getHeaders()
      }
    );

    if (!response.ok) {
      throw new Error(
        `Could not access the selected SharePoint List (${response.status}): ` +
        await response.text()
      );
    }

    return response.json() as Promise<IListInfo>;
  }

  private async getListFields(
    listId: string
  ): Promise<ISharePointField[]> {

    const url =
      `${this.webUrl}/_api/web/lists(guid'${this.escapeOData(listId)}')/fields` +
      '?$select=Title,InternalName,TypeAsString,ReadOnlyField,Hidden,' +
      'LookupList,LookupField';

    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: this.getHeaders()
      }
    );

    if (!response.ok) {
      throw new Error(
        `Could not retrieve SharePoint List columns (${response.status}): ` +
        await response.text()
      );
    }

    const data = await response.json() as {
      value?: ISharePointField[];
    };

    return data.value || [];
  }

  private async createColumn(
    listId: string,
    displayName: string,
    internalName: string,
    fieldType: string
  ): Promise<void> {

    const schemaXml =
      `<Field Type="${fieldType}" ` +
      `DisplayName="${this.escapeXml(displayName)}" ` +
      `Name="${this.escapeXml(internalName)}" ` +
      `StaticName="${this.escapeXml(internalName)}" ` +
      `Required="FALSE" ` +
      `Group="Custom Form Fields" />`;

    const endpoint =
      `${this.webUrl}/_api/web/lists(guid'${this.escapeOData(listId)}')` +
      '/fields/CreateFieldAsXml';

    const response = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({
          parameters: {
            SchemaXml: schemaXml,
            Options: 0
          }
        })
      }
    );

    if (!response.ok) {
      const errorText = await response.text();

      throw new Error(
        `Could not create the "${displayName}" column ` +
        `(${response.status}): ${errorText}`
      );
    }
  }

  private async ensureColumns(
    listId: string,
    values: { [key: string]: unknown }
  ): Promise<ISharePointField[]> {

    let fields = await this.getListFields(listId);

    for (const fieldName of Object.keys(values)) {
      if (!fieldName || fieldName === '__metadata') {
        continue;
      }

      if (this.findField(fields, fieldName)) {
        continue;
      }

      if (this.isEmpty(values[fieldName])) {
        continue;
      }

      await this.createColumn(
        listId,
        fieldName,
        this.getInternalName(fieldName),
        this.getFieldType(
          fieldName,
          values[fieldName]
        )
      );

      fields = await this.getListFields(listId);
    }

    return fields;
  }

  // ==================================================
  // SAVE SUBMISSION
  // ==================================================

  public async save(
    listId: string,
    _schemaFileServerRelativeUrl: string,
    formTitle: string,
    values: { [key: string]: unknown },
    status: string = 'submitted'
  ): Promise<ISubmissionResult> {

    if (!listId) {
      throw new Error('Please select a SharePoint List.');
    }

    const stamp = new Date()
      .toISOString()
      .replace(/[-:.TZ]/g, '')
      .substring(0, 14);

    const random = Math.random()
      .toString(36)
      .substring(2, 6)
      .toUpperCase();

    const reference =
      `${status.toLowerCase() === 'draft' ? 'DRAFT' : 'SUB'}-${stamp}-${random}`;

    // Extract files.
    const extracted = this.extractFiles(values);

    // Submission metadata.
    const submissionValues: {
      [key: string]: unknown;
    } = {
      ...extracted.clean,
      SubmissionReference: reference,
      SubmissionStatus: status,
      SubmittedOn: new Date().toISOString(),
      SubmittedBy:
        this.context.pageContext.user.displayName || ''
    };

    // Create missing columns.
    const fields = await this.ensureColumns(
      listId,
      submissionValues
    );

    const listInfo = await this.getListInfo(listId);

    const itemPayload: {
      [key: string]: unknown;
    } = {};

    // Convert form values.
    for (const fieldName of Object.keys(submissionValues)) {
      if (fieldName === '__metadata') {
        continue;
      }

      const matchingField = this.findField(
        fields,
        fieldName
      );

      if (
        !matchingField ||
        matchingField.ReadOnlyField ||
        matchingField.Hidden
      ) {
        continue;
      }

      try {
        const result = await this.buildFieldValue(
          matchingField,
          submissionValues[fieldName]
        );

        if (result) {
          itemPayload[result.name] = result.value;
        }
      } catch (error) {
        console.warn(
          `Skipped "${fieldName}" because its value could not be converted:`,
          error
        );
      }
    }

    // Default Title.
    if (
      !Object.prototype.hasOwnProperty.call(
        itemPayload,
        'Title'
      )
    ) {
      itemPayload.Title =
        formTitle || 'Form Submission';
    }

    const itemEndpoint =
      `${this.webUrl}/_api/web/lists(guid'${this.escapeOData(listId)}')/items`;

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata'
      },
      body: JSON.stringify(itemPayload)
    };

    const response = await this.context.spHttpClient.post(
      itemEndpoint,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      const errorText = await response.text();

      console.error(
        'SharePoint list item creation failed:',
        {
          status: response.status,
          payload: itemPayload,
          error: errorText
        }
      );

      throw new Error(
        `Could not save the SharePoint List Item ` +
        `(${response.status}): ${errorText}`
      );
    }

    const savedItem =
      await response.json() as IListItemCreateResponse;

    const itemId =
      (savedItem.d && savedItem.d.Id) ||
      savedItem.Id;

    if (!itemId) {
      throw new Error(
        'The item was saved, but SharePoint did not return its item ID.'
      );
    }

    // ==================================================
    // UPLOAD ATTACHMENTS
    // ==================================================

    let attachmentsUploaded = 0;
    const failedAttachments: string[] = [];

    if (extracted.files.length > 0) {
      if (listInfo.EnableAttachments !== true) {
        extracted.files.forEach(f => {
          failedAttachments.push(f.name);
        });
      } else {
        const used: {
          [name: string]: boolean;
        } = {};

        for (const pending of extracted.files) {
          const name = this.uniqueName(
            pending.name,
            used
          );

          try {
            await this.uploadAttachment(
              listId,
              itemId,
              name,
              pending.file
            );

            attachmentsUploaded++;
          } catch (error) {
            console.warn(
              `Could not attach "${name}":`,
              error
            );

            failedAttachments.push(pending.name);
          }
        }
      }
    }

    // ==================================================
    // SHAREPOINT ITEM URL
    // ==================================================

    const listServerRelativeUrl =
      listInfo.RootFolder &&
      listInfo.RootFolder.ServerRelativeUrl;

    const finalItemUrl = listServerRelativeUrl
      ? `${window.location.origin}${listServerRelativeUrl}/DispForm.aspx?ID=${itemId}`
      : `${this.webUrl}/Lists/${encodeURIComponent(listInfo.Title)}/DispForm.aspx?ID=${itemId}`;

    // ==================================================
    // SEND EMAIL
    // ==================================================

    let emailSent = false;
    let emailError = '';

    if (status.toLowerCase() !== 'draft') {
      try {
        console.log(
          'Starting submission email process...'
        );

        await this.sendSubmissionEmail(
          submissionValues,
          reference,
          itemId,
          finalItemUrl,
          formTitle
        );

        emailSent = true;

        console.log(
          `Submission email sent to ${this.testEmailRecipient}`
        );
      } catch (error) {
        emailError = this.getEmailErrorMessage(error);

        console.error(
          'Submission saved, but email sending failed:',
          error
        );

        console.error(
          'Microsoft Graph error details:',
          emailError
        );
      }
    }

    // ==================================================
    // RETURN RESULT
    // ==================================================

    return {
      reference,
      itemId,
      itemUrl: finalItemUrl,
      attachmentsUploaded,
      failedAttachments,
      emailSent,
      emailError: emailError || undefined
    };
  }
}
