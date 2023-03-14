import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDepartPivotProps {
  description: string;
  propertyToggle:boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:any;
}
