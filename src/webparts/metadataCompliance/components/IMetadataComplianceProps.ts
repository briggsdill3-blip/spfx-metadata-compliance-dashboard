import { SPFI } from '@pnp/sp';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IMetadataComplianceProps {
  theme: IReadonlyTheme | undefined;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  lockedLibrary: string;
  goodThreshold: number;
  warnThreshold: number;
  excludedFields: string[];
}