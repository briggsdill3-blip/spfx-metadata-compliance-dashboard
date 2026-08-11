import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import SiteEntryChipInput, { ISiteEntryChipInputProps } from './SiteEntryChipInput';
import { ISiteEntry } from '../components/ISiteEntry';

export interface ISiteEntryChipInputFieldProps {
  label: string;
  entries: ISiteEntry[];
  onPropertyChange: (propertyPath: string, newValue: ISiteEntry[]) => void;
}

export class PropertyPaneSiteEntryChipInputField implements IPropertyPaneField<ISiteEntryChipInputFieldProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: ISiteEntryChipInputFieldProps;
  private elem: HTMLElement | undefined;

  constructor(targetProperty: string, properties: ISiteEntryChipInputFieldProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: targetProperty,
      label: properties.label,
      entries: properties.entries,
      onPropertyChange: properties.onPropertyChange,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    } as ISiteEntryChipInputFieldProps;
  }

  public render(): void {
    if (!this.elem) {
      return;
    }
    this.onRender(this.elem);
  }

  private onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element: React.ReactElement<ISiteEntryChipInputProps> = React.createElement(SiteEntryChipInput, {
      label: this.properties.label,
      entries: this.properties.entries,
      onChange: this.onChanged.bind(this)
    });

    ReactDom.render(element, elem);
  }

  private onChanged(entries: ISiteEntry[]): void {
    this.properties.entries = entries;
    this.properties.onPropertyChange(this.targetProperty, entries);
  }
}