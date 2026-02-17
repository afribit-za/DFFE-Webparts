import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneLabel,
  PropertyPaneDropdown,
  type IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import CarouselWidgetIcons from './components/CarouselWidgetIcons';
import { ICarouselWidgetIconsProps } from './components/ICarouselWidgetIconsProps';
import { ICarouselCardConfig, ISocialIconConfig } from './models/IModels';

const DEFAULT_SITE_URL: string = 'https://afribitholdings.sharepoint.com/sites/DFFECentral';

export interface ICarouselWidgetIconsWebPartProps {
  // Card 1
  card1SiteUrl: string;
  card1ListName: string;
  card1Label: string;
  card1TitleColumn: string;
  card1ContentColumn: string;
  card1ImageColumn: string;
  card1AttachmentColumn: string;
  card1DateColumn: string;
  // Card 2
  card2SiteUrl: string;
  card2ListName: string;
  card2Label: string;
  card2TitleColumn: string;
  card2ContentColumn: string;
  card2ImageColumn: string;
  card2AttachmentColumn: string;
  card2DateColumn: string;
  // Card 3
  card3SiteUrl: string;
  card3ListName: string;
  card3Label: string;
  card3TitleColumn: string;
  card3ContentColumn: string;
  card3ImageColumn: string;
  card3AttachmentColumn: string;
  card3DateColumn: string;
  // Weather
  numberOfCities: number;
  city1Name: string;
  city2Name: string;
  city3Name: string;
  city4Name: string;
  city5Name: string;
  // Social Icons
  numberOfIcons: number;
  icon1ImageUrl: string;
  icon1LinkUrl: string;
  icon1Tooltip: string;
  icon2ImageUrl: string;
  icon2LinkUrl: string;
  icon2Tooltip: string;
  icon3ImageUrl: string;
  icon3LinkUrl: string;
  icon3Tooltip: string;
  icon4ImageUrl: string;
  icon4LinkUrl: string;
  icon4Tooltip: string;
  icon5ImageUrl: string;
  icon5LinkUrl: string;
  icon5Tooltip: string;
  icon6ImageUrl: string;
  icon6LinkUrl: string;
  icon6Tooltip: string;
}

export default class CarouselWidgetIconsWebPart extends BaseClientSideWebPart<ICarouselWidgetIconsWebPartProps> {

  private _card1ListOptions: IPropertyPaneDropdownOption[] = [];
  private _card2ListOptions: IPropertyPaneDropdownOption[] = [];
  private _card3ListOptions: IPropertyPaneDropdownOption[] = [];
  private _card1ColumnOptions: IPropertyPaneDropdownOption[] = [];
  private _card2ColumnOptions: IPropertyPaneDropdownOption[] = [];
  private _card3ColumnOptions: IPropertyPaneDropdownOption[] = [];
  private _listsLoaded: boolean = false;

  public render(): void {
    const carouselCards: ICarouselCardConfig[] = [
      {
        siteUrl: this.properties.card1SiteUrl || DEFAULT_SITE_URL,
        listName: this.properties.card1ListName || '',
        cardLabel: this.properties.card1Label || 'Updates',
        titleColumn: this.properties.card1TitleColumn || 'Title',
        contentColumn: this.properties.card1ContentColumn || '',
        imageColumn: this.properties.card1ImageColumn || '',
        attachmentColumn: this.properties.card1AttachmentColumn || '',
        dateColumn: this.properties.card1DateColumn || 'Created'
      },
      {
        siteUrl: this.properties.card2SiteUrl || DEFAULT_SITE_URL,
        listName: this.properties.card2ListName || '',
        cardLabel: this.properties.card2Label || 'News',
        titleColumn: this.properties.card2TitleColumn || 'Title',
        contentColumn: this.properties.card2ContentColumn || '',
        imageColumn: this.properties.card2ImageColumn || '',
        attachmentColumn: this.properties.card2AttachmentColumn || '',
        dateColumn: this.properties.card2DateColumn || 'Created'
      },
      {
        siteUrl: this.properties.card3SiteUrl || DEFAULT_SITE_URL,
        listName: this.properties.card3ListName || '',
        cardLabel: this.properties.card3Label || 'Announcements',
        titleColumn: this.properties.card3TitleColumn || 'Title',
        contentColumn: this.properties.card3ContentColumn || '',
        imageColumn: this.properties.card3ImageColumn || '',
        attachmentColumn: this.properties.card3AttachmentColumn || '',
        dateColumn: this.properties.card3DateColumn || 'Created'
      }
    ];

    const numberOfCities: number = this.properties.numberOfCities || 2;
    const allCityNames: string[] = [
      this.properties.city1Name,
      this.properties.city2Name,
      this.properties.city3Name,
      this.properties.city4Name,
      this.properties.city5Name
    ];
    const cities: string[] = allCityNames.slice(0, numberOfCities).filter(Boolean);

    const numberOfIcons: number = this.properties.numberOfIcons || 0;
    const socialIcons: ISocialIconConfig[] = [];
    const iconProps: Array<{ img: string; link: string; tip: string }> = [
      { img: this.properties.icon1ImageUrl, link: this.properties.icon1LinkUrl, tip: this.properties.icon1Tooltip },
      { img: this.properties.icon2ImageUrl, link: this.properties.icon2LinkUrl, tip: this.properties.icon2Tooltip },
      { img: this.properties.icon3ImageUrl, link: this.properties.icon3LinkUrl, tip: this.properties.icon3Tooltip },
      { img: this.properties.icon4ImageUrl, link: this.properties.icon4LinkUrl, tip: this.properties.icon4Tooltip },
      { img: this.properties.icon5ImageUrl, link: this.properties.icon5LinkUrl, tip: this.properties.icon5Tooltip },
      { img: this.properties.icon6ImageUrl, link: this.properties.icon6LinkUrl, tip: this.properties.icon6Tooltip }
    ];

    for (let i = 0; i < numberOfIcons; i++) {
      if (iconProps[i] && iconProps[i].img) {
        socialIcons.push({
          imageUrl: iconProps[i].img || '',
          linkUrl: iconProps[i].link || '#',
          tooltip: iconProps[i].tip || ''
        });
      }
    }

    const element: React.ReactElement<ICarouselWidgetIconsProps> = React.createElement(
      CarouselWidgetIcons,
      {
        context: this.context,
        carouselCards: carouselCards,
        cities: cities,
        socialIcons: socialIcons
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Fetches available lists from the given SharePoint site URL.
   */
  private async _fetchLists(siteUrl: string): Promise<IPropertyPaneDropdownOption[]> {
    const url: string = (siteUrl || DEFAULT_SITE_URL).replace(/\/+$/, '');
    const endpoint: string = `${url}/_api/web/lists?$filter=Hidden eq false&$select=Title,Id,BaseTemplate&$orderby=Title`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' } }
      );

      if (!response.ok) {
        return [{ key: '', text: '(Unable to load lists)' }];
      }

      const data: { value: Array<{ Title: string; Id: string }> } = await response.json();
      const options: IPropertyPaneDropdownOption[] = [
        { key: '', text: '-- Select a list --' }
      ];
      data.value.forEach((list) => {
        options.push({ key: list.Title, text: list.Title });
      });
      return options;
    } catch (error) {
      return [{ key: '', text: '(Error loading lists)' }];
    }
  }

  /**
   * Fetches available columns (fields) from a specific SharePoint list.
   * Filters out hidden/internal fields and returns user-friendly column names.
   */
  private async _fetchColumns(siteUrl: string, listName: string): Promise<IPropertyPaneDropdownOption[]> {
    if (!listName) {
      return [{ key: '', text: '-- Select a list first --' }];
    }

    const url: string = (siteUrl || DEFAULT_SITE_URL).replace(/\/+$/, '');
    const endpoint: string =
      `${url}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/fields` +
      `?$filter=Hidden eq false` +
      `&$select=InternalName,Title,TypeAsString,ReadOnlyField` +
      `&$orderby=Title`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' } }
      );

      if (!response.ok) {
        return [{ key: '', text: '(Unable to load columns)' }];
      }

      const data: { value: Array<{ InternalName: string; Title: string; TypeAsString: string }> } = await response.json();
      const options: IPropertyPaneDropdownOption[] = [
        { key: '', text: '-- Select a column --' }
      ];

      // Filter out common system fields that aren't useful
      const excludeFields: string[] = [
        'ContentType', 'Compliance_x0020_Asset_x0020_Id', '_ComplianceFlags',
        '_ComplianceTag', '_ComplianceTagWrittenTime', '_ComplianceTagUserId',
        '_IsRecord', 'Edit', 'DocIcon', 'ItemChildCount', 'FolderChildCount',
        '_UIVersionString', 'AppAuthor', 'AppEditor'
      ];

      data.value.forEach((field) => {
        if (excludeFields.indexOf(field.InternalName) === -1) {
          const displayText: string = field.Title !== field.InternalName
            ? `${field.Title} (${field.InternalName}) — ${field.TypeAsString}`
            : `${field.InternalName} — ${field.TypeAsString}`;
          options.push({ key: field.InternalName, text: displayText });
        }
      });

      return options;
    } catch (error) {
      return [{ key: '', text: '(Error loading columns)' }];
    }
  }

  /**
   * Loads list options and column options for all 3 cards when the property pane opens.
   */
  private async _loadAllListOptions(): Promise<void> {
    const [opts1, opts2, opts3] = await Promise.all([
      this._fetchLists(this.properties.card1SiteUrl),
      this._fetchLists(this.properties.card2SiteUrl),
      this._fetchLists(this.properties.card3SiteUrl)
    ]);
    this._card1ListOptions = opts1;
    this._card2ListOptions = opts2;
    this._card3ListOptions = opts3;

    // Also load columns for any already-selected lists
    const [cols1, cols2, cols3] = await Promise.all([
      this._fetchColumns(this.properties.card1SiteUrl, this.properties.card1ListName),
      this._fetchColumns(this.properties.card2SiteUrl, this.properties.card2ListName),
      this._fetchColumns(this.properties.card3SiteUrl, this.properties.card3ListName)
    ]);
    this._card1ColumnOptions = cols1;
    this._card2ColumnOptions = cols2;
    this._card3ColumnOptions = cols3;

    this._listsLoaded = true;
    this.context.propertyPane.refresh();
  }

  /**
   * Called when the property pane is opened — trigger list loading.
   */
  protected onPropertyPaneConfigurationStart(): void {
    if (!this._listsLoaded) {
      this._loadAllListOptions().catch(function () { /* noop */ });
    }
  }

  /**
   * Called when a property pane field changes — reload lists if site URL changed.
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    // Card 1: Site URL changed → reload lists + clear columns
    if (propertyPath === 'card1SiteUrl' && oldValue !== newValue) {
      this._card1ListOptions = [{ key: '', text: 'Loading lists...' }];
      this._card1ColumnOptions = [{ key: '', text: '-- Select a list first --' }];
      this.properties.card1ListName = '';
      this.context.propertyPane.refresh();
      this._fetchLists(newValue).then((opts) => {
        this._card1ListOptions = opts;
        this.context.propertyPane.refresh();
      }).catch(function () { /* noop */ });
    }
    // Card 1: List changed → reload columns
    if (propertyPath === 'card1ListName' && oldValue !== newValue) {
      this._card1ColumnOptions = [{ key: '', text: 'Loading columns...' }];
      this.context.propertyPane.refresh();
      this._fetchColumns(this.properties.card1SiteUrl, newValue).then((opts) => {
        this._card1ColumnOptions = opts;
        this.context.propertyPane.refresh();
      }).catch(function () { /* noop */ });
    }

    // Card 2: Site URL changed → reload lists + clear columns
    if (propertyPath === 'card2SiteUrl' && oldValue !== newValue) {
      this._card2ListOptions = [{ key: '', text: 'Loading lists...' }];
      this._card2ColumnOptions = [{ key: '', text: '-- Select a list first --' }];
      this.properties.card2ListName = '';
      this.context.propertyPane.refresh();
      this._fetchLists(newValue).then((opts) => {
        this._card2ListOptions = opts;
        this.context.propertyPane.refresh();
      }).catch(function () { /* noop */ });
    }
    // Card 2: List changed → reload columns
    if (propertyPath === 'card2ListName' && oldValue !== newValue) {
      this._card2ColumnOptions = [{ key: '', text: 'Loading columns...' }];
      this.context.propertyPane.refresh();
      this._fetchColumns(this.properties.card2SiteUrl, newValue).then((opts) => {
        this._card2ColumnOptions = opts;
        this.context.propertyPane.refresh();
      }).catch(function () { /* noop */ });
    }

    // Card 3: Site URL changed → reload lists + clear columns
    if (propertyPath === 'card3SiteUrl' && oldValue !== newValue) {
      this._card3ListOptions = [{ key: '', text: 'Loading lists...' }];
      this._card3ColumnOptions = [{ key: '', text: '-- Select a list first --' }];
      this.properties.card3ListName = '';
      this.context.propertyPane.refresh();
      this._fetchLists(newValue).then((opts) => {
        this._card3ListOptions = opts;
        this.context.propertyPane.refresh();
      }).catch(function () { /* noop */ });
    }
    // Card 3: List changed → reload columns
    if (propertyPath === 'card3ListName' && oldValue !== newValue) {
      this._card3ColumnOptions = [{ key: '', text: 'Loading columns...' }];
      this.context.propertyPane.refresh();
      this._fetchColumns(this.properties.card3SiteUrl, newValue).then((opts) => {
        this._card3ColumnOptions = opts;
        this.context.propertyPane.refresh();
      }).catch(function () { /* noop */ });
    }
  }

  /**
   * Creates a custom property pane field with file upload + URL input for social icons.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _createIconUploadField(targetProperty: string, label: string): any {
    const self = this;
    return {
      type: 1, // PropertyPaneFieldType.Custom
      targetProperty: targetProperty,
      properties: {
        key: `iconUpload_${targetProperty}`,
        onRender: (elem: HTMLElement): void => {
          const currentValue: string = (self.properties as any)[targetProperty] || ''; // eslint-disable-line @typescript-eslint/no-explicit-any
          elem.innerHTML = '';

          // Container
          const container: HTMLDivElement = document.createElement('div');
          container.style.cssText = 'margin-bottom:10px;';

          // Label
          const labelEl: HTMLLabelElement = document.createElement('label');
          labelEl.textContent = label;
          labelEl.style.cssText = 'display:block;font-weight:600;font-size:14px;margin-bottom:5px;color:#323130;font-family:"Segoe UI",Arial,sans-serif;';
          container.appendChild(labelEl);

          // Row: URL input + upload button
          const row: HTMLDivElement = document.createElement('div');
          row.style.cssText = 'display:flex;gap:6px;align-items:center;';

          const urlInput: HTMLInputElement = document.createElement('input');
          urlInput.type = 'text';
          urlInput.value = (currentValue.indexOf('data:') === 0) ? '(uploaded file)' : currentValue;
          urlInput.placeholder = 'Paste URL or upload file';
          urlInput.style.cssText = 'flex:1;padding:5px 8px;border:1px solid #8a8886;border-radius:2px;font-size:13px;font-family:"Segoe UI",Arial,sans-serif;outline:none;box-sizing:border-box;';
          urlInput.onfocus = (): void => { urlInput.style.borderColor = '#0078d4'; };
          urlInput.onblur = (): void => {
            urlInput.style.borderColor = '#8a8886';
            const newVal: string = urlInput.value;
            if (newVal !== currentValue && newVal !== '(uploaded file)') {
              (self.properties as any)[targetProperty] = newVal; // eslint-disable-line @typescript-eslint/no-explicit-any
              self.render();
            }
          };
          row.appendChild(urlInput);

          // Upload button
          const uploadBtn: HTMLButtonElement = document.createElement('button');
          uploadBtn.textContent = 'Upload';
          uploadBtn.type = 'button';
          uploadBtn.style.cssText = 'padding:5px 12px;border:1px solid #0078d4;background:#0078d4;color:#fff;border-radius:2px;cursor:pointer;font-size:12px;font-family:"Segoe UI",Arial,sans-serif;white-space:nowrap;';

          const fileInput: HTMLInputElement = document.createElement('input');
          fileInput.type = 'file';
          fileInput.accept = 'image/png,image/svg+xml,image/jpeg,image/gif,image/webp';
          fileInput.style.display = 'none';

          uploadBtn.onclick = (): void => { fileInput.click(); };

          fileInput.onchange = (): void => {
            const file: File | undefined = fileInput.files ? fileInput.files[0] : undefined;
            if (file) {
              const reader: FileReader = new FileReader();
              reader.onload = (ev: ProgressEvent<FileReader>): void => {
                const dataUrl: string = (ev.target as FileReader).result as string;
                (self.properties as any)[targetProperty] = dataUrl; // eslint-disable-line @typescript-eslint/no-explicit-any
                urlInput.value = '(uploaded file)';
                self.render();
                self.context.propertyPane.refresh();
              };
              reader.readAsDataURL(file);
            }
          };

          row.appendChild(uploadBtn);
          row.appendChild(fileInput);
          container.appendChild(row);

          // Preview thumbnail
          if (currentValue) {
            const preview: HTMLImageElement = document.createElement('img');
            preview.src = currentValue;
            preview.alt = 'Icon preview';
            preview.style.cssText = 'width:28px;height:28px;margin-top:6px;border-radius:4px;object-fit:contain;border:1px solid #edebe9;background:#faf9f8;padding:2px;';
            container.appendChild(preview);
          }

          elem.appendChild(container);
        },
        onDispose: (elem: HTMLElement): void => {
          elem.innerHTML = '';
        }
      }
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const numberOfCities: number = this.properties.numberOfCities || 2;
    const numberOfIcons: number = this.properties.numberOfIcons || 0;

    // Build dynamic city fields
    const cityFields: any[] = [
      PropertyPaneSlider('numberOfCities', { label: 'Number of Cities', min: 2, max: 5, step: 1 }),
      PropertyPaneLabel('weatherInfo', { text: 'Weather data provided by Open-Meteo (free, no API key needed).' })
    ];
    const cityNames: string[] = ['city1Name', 'city2Name', 'city3Name', 'city4Name', 'city5Name'];
    const cityLabels: string[] = ['City 1', 'City 2', 'City 3', 'City 4', 'City 5'];
    for (let i = 0; i < numberOfCities; i++) {
      cityFields.push(
        PropertyPaneTextField(cityNames[i], { label: cityLabels[i] })
      );
    }

    // Build dynamic icon fields
    const iconFields: any[] = [
      PropertyPaneSlider('numberOfIcons', { label: 'Number of Social Icons', min: 0, max: 6, step: 1 })
    ];
    for (let i = 0; i < numberOfIcons; i++) {
      const n: number = i + 1;
      iconFields.push(
        PropertyPaneLabel(`iconLabel${n}`, { text: `--- Icon ${n} ---` }),
        this._createIconUploadField(`icon${n}ImageUrl`, `Icon ${n} Image`),
        PropertyPaneTextField(`icon${n}LinkUrl`, { label: `Icon ${n} Link URL` }),
        PropertyPaneTextField(`icon${n}Tooltip`, { label: `Icon ${n} Tooltip` })
      );
    }

    // Loading placeholder for list/column dropdowns
    const loadingOptions: IPropertyPaneDropdownOption[] = [{ key: '', text: 'Loading lists...' }];
    const loadingColumnOptions: IPropertyPaneDropdownOption[] = [{ key: '', text: '-- Select a list first --' }];

    return {
      pages: [
        // Page 1: Carousel Card 1
        {
          header: { description: 'Card 1 — Updates Configuration' },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneTextField('card1SiteUrl', {
                  label: 'Site URL',
                  placeholder: DEFAULT_SITE_URL,
                  description: `Default: ${DEFAULT_SITE_URL}`
                }),
                PropertyPaneDropdown('card1ListName', {
                  label: 'Select List',
                  options: this._card1ListOptions.length > 0 ? this._card1ListOptions : loadingOptions
                }),
                PropertyPaneTextField('card1Label', { label: 'Card Label (e.g. Updates)' })
              ]
            },
            {
              groupName: 'Column Mapping',
              groupFields: [
                PropertyPaneDropdown('card1TitleColumn', {
                  label: 'Title Column',
                  options: this._card1ColumnOptions.length > 0 ? this._card1ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card1ContentColumn', {
                  label: 'Content Column',
                  options: this._card1ColumnOptions.length > 0 ? this._card1ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card1ImageColumn', {
                  label: 'Image Column',
                  options: this._card1ColumnOptions.length > 0 ? this._card1ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card1AttachmentColumn', {
                  label: 'Attachment Column',
                  options: this._card1ColumnOptions.length > 0 ? this._card1ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card1DateColumn', {
                  label: 'Date Uploaded Column',
                  options: this._card1ColumnOptions.length > 0 ? this._card1ColumnOptions : loadingColumnOptions
                })
              ]
            }
          ]
        },
        // Page 2: Carousel Card 2
        {
          header: { description: 'Card 2 — News Configuration' },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneTextField('card2SiteUrl', {
                  label: 'Site URL',
                  placeholder: DEFAULT_SITE_URL,
                  description: `Default: ${DEFAULT_SITE_URL}`
                }),
                PropertyPaneDropdown('card2ListName', {
                  label: 'Select List',
                  options: this._card2ListOptions.length > 0 ? this._card2ListOptions : loadingOptions
                }),
                PropertyPaneTextField('card2Label', { label: 'Card Label (e.g. News)' })
              ]
            },
            {
              groupName: 'Column Mapping',
              groupFields: [
                PropertyPaneDropdown('card2TitleColumn', {
                  label: 'Title Column',
                  options: this._card2ColumnOptions.length > 0 ? this._card2ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card2ContentColumn', {
                  label: 'Content Column',
                  options: this._card2ColumnOptions.length > 0 ? this._card2ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card2ImageColumn', {
                  label: 'Image Column',
                  options: this._card2ColumnOptions.length > 0 ? this._card2ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card2AttachmentColumn', {
                  label: 'Attachment Column',
                  options: this._card2ColumnOptions.length > 0 ? this._card2ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card2DateColumn', {
                  label: 'Date Uploaded Column',
                  options: this._card2ColumnOptions.length > 0 ? this._card2ColumnOptions : loadingColumnOptions
                })
              ]
            }
          ]
        },
        // Page 3: Carousel Card 3
        {
          header: { description: 'Card 3 — Announcements Configuration' },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneTextField('card3SiteUrl', {
                  label: 'Site URL',
                  placeholder: DEFAULT_SITE_URL,
                  description: `Default: ${DEFAULT_SITE_URL}`
                }),
                PropertyPaneDropdown('card3ListName', {
                  label: 'Select List',
                  options: this._card3ListOptions.length > 0 ? this._card3ListOptions : loadingOptions
                }),
                PropertyPaneTextField('card3Label', { label: 'Card Label (e.g. Announcements)' })
              ]
            },
            {
              groupName: 'Column Mapping',
              groupFields: [
                PropertyPaneDropdown('card3TitleColumn', {
                  label: 'Title Column',
                  options: this._card3ColumnOptions.length > 0 ? this._card3ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card3ContentColumn', {
                  label: 'Content Column',
                  options: this._card3ColumnOptions.length > 0 ? this._card3ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card3ImageColumn', {
                  label: 'Image Column',
                  options: this._card3ColumnOptions.length > 0 ? this._card3ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card3AttachmentColumn', {
                  label: 'Attachment Column',
                  options: this._card3ColumnOptions.length > 0 ? this._card3ColumnOptions : loadingColumnOptions
                }),
                PropertyPaneDropdown('card3DateColumn', {
                  label: 'Date Uploaded Column',
                  options: this._card3ColumnOptions.length > 0 ? this._card3ColumnOptions : loadingColumnOptions
                })
              ]
            }
          ]
        },
        // Page 4: Weather Configuration
        {
          header: { description: 'Weather Widget — powered by Open-Meteo (free, no API key needed)' },
          groups: [
            {
              groupName: 'Weather Settings',
              groupFields: cityFields
            }
          ]
        },
        // Page 5: Social Icons Configuration
        {
          header: { description: 'Social Media Icons Configuration' },
          groups: [
            {
              groupName: 'Social Icons',
              groupFields: iconFields
            }
          ]
        }
      ]
    };
  }
}
