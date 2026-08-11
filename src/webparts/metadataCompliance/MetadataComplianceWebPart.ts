import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { spfi, SPFx as spSPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

import * as strings from 'MetadataComplianceWebPartStrings';
import MetadataCompliance from './components/MetadataCompliance';
import { IMetadataComplianceProps } from './components/IMetadataComplianceProps';
import { ISiteEntry } from './components/ISiteEntry';
import { PropertyPaneSiteEntryChipInputField } from './controls/SiteEntryChipInputField';
import { PropertyPaneLibraryExclusionChecklistField } from './controls/LibraryExclusionChecklistField';

export interface IMetadataComplianceWebPartProps {
  targetSites: ISiteEntry[];
  excludedLibraries: string[];
  goodThreshold: string;
  warnThreshold: string;
  excludedFields: string;
}

const parseThreshold = (raw: string, fallback: number): number => {
  const parsed = parseInt(raw, 10);
  if (isNaN(parsed) || parsed < 0 || parsed > 100) {
    return fallback;
  }
  return parsed;
};

const parseExcludedFields = (raw: string): string[] => {
  if (!raw) {
    return [];
  }
  return raw.split(',').map(f => f.trim().toLowerCase()).filter(f => f.length > 0);
};

export default class MetadataComplianceWebPart extends BaseClientSideWebPart<IMetadataComplianceWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _environmentMessage: string = '';
  private _sp!: SPFI;
  private _targetSitesField: PropertyPaneSiteEntryChipInputField | undefined;
  private _exclusionChecklistField: PropertyPaneLibraryExclusionChecklistField | undefined;

  public render(): void {
    if (!this._sp) {
      return;
    }

    const element: React.ReactElement<IMetadataComplianceProps> = React.createElement(
      MetadataCompliance,
      {
        theme: this._theme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        sp: this._sp,
        targetSites: this.properties.targetSites || [],
        excludedLibraries: this.properties.excludedLibraries || [],
        excludedFields: parseExcludedFields(this.properties.excludedFields),
        goodThreshold: parseThreshold(this.properties.goodThreshold, 90),
        warnThreshold: parseThreshold(this.properties.warnThreshold, 70)
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._sp = spfi().using(spSPFx(this.context));

    if (!this.properties.targetSites) {
      this.properties.targetSites = [];
    }
    if (!this.properties.excludedLibraries) {
      this.properties.excludedLibraries = [];
    }

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._theme = currentTheme;
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('3.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this._targetSitesField = new PropertyPaneSiteEntryChipInputField('targetSites', {
      label: 'Target Sites',
      entries: this.properties.targetSites || [],
      onPropertyChange: (propertyPath: string, newValue: ISiteEntry[]) => {
        this.properties.targetSites = newValue;
        this.render();

        if (this._targetSitesField) {
          this._targetSitesField.properties.entries = newValue;
          this._targetSitesField.render();
        }

        if (this._exclusionChecklistField) {
          this._exclusionChecklistField.properties.sites = newValue;
          this._exclusionChecklistField.render();
        }
      }
    });

    this._exclusionChecklistField = new PropertyPaneLibraryExclusionChecklistField('excludedLibraries', {
      label: 'Included Libraries',
      sp: this._sp,
      sites: this.properties.targetSites || [],
      excludedLibraries: this.properties.excludedLibraries || [],
      onPropertyChange: (propertyPath: string, newValue: string[]) => {
        this.properties.excludedLibraries = newValue;
        this.render();

        if (this._exclusionChecklistField) {
          this._exclusionChecklistField.properties.excludedLibraries = newValue;
        }
      }
    });

    const groups: IPropertyPaneGroup[] = [
      {
        groupName: 'Sites and Libraries',
        groupFields: [
          this._targetSitesField,
          this._exclusionChecklistField
        ]
      },
      {
        groupName: 'Data Source',
        groupFields: [
          PropertyPaneTextField('excludedFields', {
            label: 'Excluded Fields (comma-separated)',
            description: 'Metadata columns to ignore when calculating completeness, e.g. Notes, Comments'
          })
        ]
      },
      {
        groupName: 'Compliance Thresholds',
        groupFields: [
          PropertyPaneTextField('goodThreshold', {
            label: 'On Target threshold (%)',
            description: 'Default 90. Percent complete at or above this is shown as green.'
          }),
          PropertyPaneTextField('warnThreshold', {
            label: 'Needs Attention threshold (%)',
            description: 'Default 70. Percent complete at or above this (but below target) is shown as amber.'
          })
        ]
      }
    ];

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups
        }
      ]
    };
  }
}