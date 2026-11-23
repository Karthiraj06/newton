import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { IJsonSchema } from '../models/ISchemaDrivenFormProps';

export class SharePointSchemaReader {
  private context: WebPartContext;

  public constructor(context: WebPartContext) {
    this.context = context;
  }

  /** Reads a JSON file from SharePoint by its server-relative URL. */
  public async read(serverRelativeUrl: string): Promise<IJsonSchema> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const path = encodeURIComponent(serverRelativeUrl.replace(/'/g, "''"));

    const response = await this.context.spHttpClient.get(
      `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${path}')/$value`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(
        `Could not read the schema file (${response.status}): ${await response.text()}`
      );
    }

    // Remove a UTF-8 byte order mark if the file has one.
    const text = (await response.text()).replace(/^\uFEFF/, '');

    try {
      return JSON.parse(text) as IJsonSchema;
    } catch (error) {
      throw new Error(
        `The schema file is not valid JSON: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }
}