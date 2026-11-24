import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import { SPHttpClient } from '@microsoft/sp-http';

import SchemaDrivenForm from './components/SchemaDrivenForm';
import { ISchemaDrivenFormProps } from './models/ISchemaDrivenFormProps';

// --------------------------------------------------
// Web Part Properties
// --------------------------------------------------

export interface ISchemaDrivenFormWebPartProps {
  libraryTitle: string;
  schemaFileServerRelativeUrl: string;

  // Stores the selected SharePoint List GUID.
  // Kept with the old name for compatibility.
  submissionLibraryId: string;

  // When true, the web part covers the whole browser window in view mode.
  fullScreen: boolean;
}

// --------------------------------------------------
// Interfaces
// --------------------------------------------------

interface IListItem {
  Id: string;
  Title: string;
  BaseTemplate: number;
  Hidden: boolean;
}

interface IFileItem {
  File?: {
    Name: string;
    ServerRelativeUrl: string;
  };
  FSObjType?: number;
}

interface IResult<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

interface IOption {
  key: string;
  text: string;
}

// --------------------------------------------------
// Dynamic Form Web Part
// --------------------------------------------------

export default class DynamicFormWebPart extends BaseClientSideWebPart<ISchemaDrivenFormWebPartProps> {

  private schemaLibraries: IOption[] = [];
  private submissionLists: IOption[] = [];
  private schemaFiles: IOption[] = [];

  private loadingLibraries = false;
  private loadingFiles = false;

  // Tracks whether this web part has applied its own full-screen styles.
  // Prevents layout-affecting style writes on every render.
  private fullScreenApplied = false;

  // --------------------------------------------------
  // Initialize
  // Nothing is loaded here, so visitors with view-only
  // access do not trigger library/list REST calls.
  // Dropdown data loads when the property pane opens.
  // --------------------------------------------------

  public async onInit(): Promise<void> {
    await super.onInit();
  }

  // --------------------------------------------------
  // Load property pane data only for editors
  // --------------------------------------------------

  protected onPropertyPaneConfigurationStart(): void {
    this.loadLibraries()
      .then(async () => {
        if (this.properties.libraryTitle) {
          await this.loadSchemaFiles(this.properties.libraryTitle);
        }
      })
      .catch((error: unknown) => {
        console.error('Could not load property pane data:', error);
      });
  }

  // --------------------------------------------------
  // Full Screen
  // Applies only in view mode so the property pane
  // stays reachable while editing.
  // Deferred one frame to avoid ResizeObserver noise.
  // --------------------------------------------------

  private applyFullScreen(): void {
    const el = this.domElement;

    if (!el) {
      return;
    }

    // Never apply full-screen positioning while the page is being edited.
    const shouldApply =
      this.properties.fullScreen === true &&
      this.displayMode === DisplayMode.Read;

    // Avoid repeatedly changing width, height and position during render.
    if (shouldApply === this.fullScreenApplied) {
      return;
    }

    this.fullScreenApplied = shouldApply;

    if (shouldApply) {
      el.classList.add('sdfFullScreen');
      el.style.position = 'fixed';
      el.style.top = '0';
      el.style.left = '0';
      el.style.width = '100vw';
      el.style.height = '100vh';
      el.style.zIndex = '1000';
      el.style.overflow = 'auto';
      el.style.backgroundColor = '#ffffff';
    } else {
      el.classList.remove('sdfFullScreen');

      // Remove only styles owned by this method; do not clear cssText.
      el.style.removeProperty('position');
      el.style.removeProperty('top');
      el.style.removeProperty('left');
      el.style.removeProperty('width');
      el.style.removeProperty('height');
      el.style.removeProperty('z-index');
      el.style.removeProperty('overflow');
      el.style.removeProperty('background-color');
    }
  }

  // --------------------------------------------------
  // Read URL Query Parameters
  // --------------------------------------------------

  private getQueryValue(name: string): string {
    try {
      const params = new URLSearchParams(window.location.search);
      const wanted = name.toLowerCase();

      let found = '';

      params.forEach((value, key) => {
        if (!found && key.toLowerCase() === wanted) {
          found = value;
        }
      });

      return found.trim();
    } catch (error) {
      console.error('Unable to read URL parameter:', error);
      return '';
    }
  }

  // --------------------------------------------------
  // Get Schema URL
  // Supports absolute and server-relative URLs
  // --------------------------------------------------

  private getSchemaFromUrl(): string {
    const value = this.getQueryValue('schema');

    if (!value) {
      return '';
    }

    try {
      const schemaUrl = new URL(value, window.location.origin);

      if (schemaUrl.origin !== window.location.origin) {
        console.error('External schema URLs are not allowed.');
        return '';
      }

      if (!schemaUrl.pathname.toLowerCase().endsWith('.json')) {
        console.error('Schema URL must point to a JSON file.');
        return '';
      }

      return decodeURIComponent(schemaUrl.pathname);
    } catch (error) {
      console.error('Invalid schema URL:', error);
      return '';
    }
  }

  // --------------------------------------------------
  // Get Submission List ID From URL
  // Supports:
  // ?submissionListId=GUID
  // ?submissionLibraryId=GUID (legacy)
  // --------------------------------------------------

  private getSubmissionListFromUrl(): string {
    const value =
      this.getQueryValue('submissionListId') ||
      this.getQueryValue('submissionLibraryId');

    return /^[0-9a-fA-F-]{36}$/.test(value) ? value : '';
  }

  // --------------------------------------------------
  // Get All SharePoint API Results
  // --------------------------------------------------

  private async getAll<T>(url: string): Promise<T[]> {
    const all: T[] = [];
    let next: string | undefined = url;

    while (next) {
      const response = await this.context.spHttpClient.get(
        next,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(
          `SharePoint request failed: ${response.status} ${await response.text()}`
        );
      }

      const page = await response.json() as IResult<T>;

      all.push(...(page.value || []));

      next = page['@odata.nextLink'];
    }

    return all;
  }

  // --------------------------------------------------
  // Load Schema Document Libraries and Submission Lists
  // --------------------------------------------------

  private async loadLibraries(): Promise<void> {
    this.loadingLibraries = true;

    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;

      // Schema source: Document Libraries (BaseTemplate 101).
      const libraryUrl =
        `${webUrl}/_api/web/lists` +
        `?$select=Id,Title,BaseTemplate,Hidden` +
        `&$filter=BaseTemplate eq 101 and Hidden eq false` +
        `&$orderby=Title`;

      // Submission destination: Custom SharePoint Lists (BaseTemplate 100).
      const listUrl =
        `${webUrl}/_api/web/lists` +
        `?$select=Id,Title,BaseTemplate,Hidden` +
        `&$filter=BaseTemplate eq 100 and Hidden eq false` +
        `&$orderby=Title`;

      const [libraries, lists] = await Promise.all([
        this.getAll<IListItem>(libraryUrl),
        this.getAll<IListItem>(listUrl)
      ]);

      this.schemaLibraries = libraries.map((item) => ({
        key: item.Title,
        text: item.Title
      }));

      this.submissionLists = lists.map((item) => ({
        key: item.Id.replace(/[{}]/g, ''),
        text: item.Title
      }));

      if (
        this.properties.libraryTitle &&
        !libraries.some(
          (item) => item.Title === this.properties.libraryTitle
        )
      ) {
        this.properties.libraryTitle = '';
      }

      if (
        this.properties.submissionLibraryId &&
        !lists.some(
          (item) =>
            item.Id.replace(/[{}]/g, '').toLowerCase() ===
            this.properties.submissionLibraryId.replace(/[{}]/g, '').toLowerCase()
        )
      ) {
        this.properties.submissionLibraryId = '';
      }
    } catch (error) {
      console.error('Could not load SharePoint libraries and lists:', error);
    } finally {
      this.loadingLibraries = false;
      this.context.propertyPane.refresh();
    }
  }

  // --------------------------------------------------
  // Load JSON Files From Selected Document Library
  // --------------------------------------------------

  private async loadSchemaFiles(libraryTitle: string): Promise<void> {
    this.loadingFiles = true;
    this.schemaFiles = [];

    try {
      const title = libraryTitle.replace(/'/g, "''");

      const url =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${title}')/items` +
        `?$select=File/Name,File/ServerRelativeUrl,FSObjType` +
        `&$expand=File` +
        `&$filter=FSObjType eq 0` +
        `&$top=5000`;

      const items = await this.getAll<IFileItem>(url);

      const files: IOption[] = [];

      items.forEach((item) => {
        if (
          item.File &&
          item.File.Name.toLowerCase().endsWith('.json')
        ) {
          files.push({
            key: item.File.ServerRelativeUrl,
            text: item.File.Name
          });
        }
      });

      this.schemaFiles = files.sort(
        (a, b) => a.text.localeCompare(b.text)
      );

      if (
        !this.schemaFiles.some(
          (file) =>
            file.key === this.properties.schemaFileServerRelativeUrl
        )
      ) {
        this.properties.schemaFileServerRelativeUrl = '';
      }
    } catch (error) {
      console.error('Could not load JSON schema files:', error);
    } finally {
      this.loadingFiles = false;
      this.context.propertyPane.refresh();
    }
  }

  // --------------------------------------------------
  // Property Pane Change Handler
  // --------------------------------------------------

  protected onPropertyPaneFieldChanged(
    path: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(path, oldValue, newValue);

    if (path === 'libraryTitle' && newValue !== oldValue) {
      this.properties.schemaFileServerRelativeUrl = '';

      if (typeof newValue === 'string' && newValue) {
        this.loadSchemaFiles(newValue).catch((error: unknown) => {
          console.error('Error loading schema files:', error);
        });
      } else {
        this.schemaFiles = [];
        this.context.propertyPane.refresh();
      }
    }
  }

  // --------------------------------------------------
  // Render Web Part
  // --------------------------------------------------

  public render(): void {
    this.applyFullScreen();

    // Priority 1: URL schema
    // Priority 2: Property Pane schema

    const urlSchema = this.getSchemaFromUrl();

    const schemaUrl =
      urlSchema ||
      this.properties.schemaFileServerRelativeUrl ||
      '';

    // Priority 1: URL submission list ID
    // Priority 2: Property Pane submission list ID

    const urlSubmissionId = this.getSubmissionListFromUrl();

    const submissionId =
      urlSubmissionId ||
      this.properties.submissionLibraryId ||
      '';

    const element: React.ReactElement<ISchemaDrivenFormProps> =
      React.createElement(SchemaDrivenForm, {
        context: this.context,
        libraryTitle: this.properties.libraryTitle || '',
        schemaFileServerRelativeUrl: schemaUrl,
        submissionLibraryId: submissionId
      });

    ReactDom.render(element, this.domElement);
  }

  // --------------------------------------------------
  // Dispose
  // --------------------------------------------------

  protected onDispose(): void {
    const el = this.domElement;

    if (el) {
      el.classList.remove('sdfFullScreen');

      // Clean up only the full-screen styles owned by this web part.
      el.style.removeProperty('position');
      el.style.removeProperty('top');
      el.style.removeProperty('left');
      el.style.removeProperty('width');
      el.style.removeProperty('height');
      el.style.removeProperty('z-index');
      el.style.removeProperty('overflow');
      el.style.removeProperty('background-color');
    }

    this.fullScreenApplied = false;
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // --------------------------------------------------
  // Data Version
  // --------------------------------------------------

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // --------------------------------------------------
  // Property Pane Configuration
  // --------------------------------------------------

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              'Configure the JSON schema source and SharePoint submission list.'
          },
          groups: [
            {
              groupName: 'Schema source',
              groupFields: [
                PropertyPaneDropdown('libraryTitle', {
                  label: 'Schema Document Library',
                  options: this.schemaLibraries,
                  disabled: this.loadingLibraries
                }),

                PropertyPaneDropdown('schemaFileServerRelativeUrl', {
                  label: 'JSON Schema File',
                  options: this.schemaFiles,
                  disabled:
                    !this.properties.libraryTitle ||
                    this.loadingFiles ||
                    this.schemaFiles.length === 0
                })
              ]
            },
            {
              groupName: 'Submission destination',
              groupFields: [
                PropertyPaneDropdown('submissionLibraryId', {
                  label: 'SharePoint Submission List',
                  options: this.submissionLists,
                  disabled:
                    this.loadingLibraries ||
                    this.submissionLists.length === 0
                })
              ]
            },
            {
              groupName: 'Display',
              groupFields: [
                PropertyPaneToggle('fullScreen', {
                  label: 'Run in full screen',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}