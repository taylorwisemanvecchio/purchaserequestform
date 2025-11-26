import type { SPFI } from '@pnp/sp';

export interface IPurchaseRequestFormProps {
  sp: SPFI;                // PnPjs instance passed from the WebPart
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
