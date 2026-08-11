import { SPFI } from '@pnp/sp';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISiteEntry } from './ISiteEntry';

export interface IMetadataComplianceProps {
  theme: IReadonlyTheme | undefined;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  targetSites: ISiteEntry[];
  excludedLibraries: string[];
  excludedFields: string[];
  goodThreshold: number;
  warnThreshold: number;
}