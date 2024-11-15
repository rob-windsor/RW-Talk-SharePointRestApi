import { SPHttpClient } from '@microsoft/sp-http';

export interface IHelloRestApiProps {
  hasTeamsContext: boolean;
  domElement: HTMLElement;
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  legacyPageContext: any;
}
