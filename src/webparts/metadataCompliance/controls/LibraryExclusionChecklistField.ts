import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { SPFI } from '@pnp/sp';
import LibraryExclusionChecklist, { ILibraryExclusionChecklistProps } from './LibraryExclusionChecklist';
import { ISiteEntry } from '../components/ISiteEntry';

export interface ILibraryExclusionChecklistFieldProps {
  label: string;
  sp: SPFI;
  sites: ISiteEntry[];
  excludedLibraries: string[];
  onPropertyChange: (propertyPath: string, newValue: string[]) => void;
}

export class PropertyPaneLibraryExclusionChecklistField implements IPropertyPaneField<ILibraryExclusionChecklistFieldProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: ILibraryExclusionChecklistFieldProps;
  private elem: HTMLElement | undefined;

  constructor(targetProperty: string, properties: ILibraryExclusionChecklistFieldProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: targetProperty,
      label: properties.label,
      sp: properties.sp,
      sites: properties.sites,
      excludedLibraries: properties.excludedLibraries,
      onPropertyChange: properties.onPropertyChange,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    } as ILibraryExclusionChecklistFieldProps;
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

    const element: React.ReactElement<ILibraryExclusionChecklistProps> = React.createElement(LibraryExclusionChecklist, {
      label: this.properties.label,
      sp: this.properties.sp,
      sites: this.properties.sites,
      excludedLibraries: this.properties.excludedLibraries,
      onChange: this.onChanged.bind(this)
    });

    ReactDom.render(element, elem);
  }

  private onChanged(excludedLibraries: string[]): void {
    this.properties.excludedLibraries = excludedLibraries;
    this.properties.onPropertyChange(this.targetProperty, excludedLibraries);
  }
}