import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { ReadView } from './quickView/ReadView';
import { CreateView } from './quickView/CreateView';
import { UpdateView } from './quickView/UpdateView';
import { MessageView } from './quickView/MessageView';
import { MediumCardView } from './cardView/MediumCardView';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DialogueView } from './quickView/DialogueView';
import { ShowAllView } from './quickView/ShowAllView';
import { PeopleDetailsPropertyPane } from './PeopleDetailsPropertyPane';
import { PnPServices } from '../../Services/PnPServices';

export interface IPeopleDetailsAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
}

export interface IPeopleDetailsAdaptiveCardExtensionState {
  currentIndex: number;
  peopleData: any[];
  countryData: any[];
  messageBar: {
    text: string;
    success: boolean;
    iconUrl: string;
    color: string;
  };
  context?: any;
  success_imgPath?: any;
  error_imgPath?: any;
}

export interface IPeopleData {
  title: string;
  email: string;
  jobTitle: string;
  country: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'PeopleDetails_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'PeopleDetails_QUICK_VIEW';
export const READ_VIEW_REGISTRY_ID: string = 'PeopleDetails_READ_VIEW';
export const CREATE_VIEW_REGISTRY_ID: string = 'PeopleDetails_CREATE_VIEW';
export const UPDATE_VIEW_REGISTRY_ID: string = 'PeopleDetails_UPDATE_VIEW';
export const MESSAGE_VIEW_REGISTRY_ID: string = 'PeopleDetails_MESSAGE_VIEW';
export const DIALOGUE_VIEW_REGISTRY_ID: string = 'PeopleDetails_Dialogue_VIEW';
export const SHOWALLMEDIUM_VIEW_REGISTRY_ID: string = 'PeopleDetails_Medium_VIEW';

const MEDIUM_VIEW_REGISTRY_ID: string = 'PeopleDetails_MEDIUM_VIEW';

export default class PeopleDetailsAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IPeopleDetailsAdaptiveCardExtensionProps,
  IPeopleDetailsAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: PeopleDetailsPropertyPane | undefined;

  public onInit = async () => {
    sp.setup({
      spfxContext: this.context
    });
    let refreshData: any = await PnPServices.refreshData();
    this.state = {
      currentIndex: 0,
      peopleData: refreshData["peopleData"],
      countryData: refreshData["countryData"],
      messageBar: {
        text: "",
        success: true,
        iconUrl: "",
        color: ""
      },
      context: this.context.pageContext,
      success_imgPath: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiA/PjxzdmcgaWQ9IkxheWVyXzEiIHN0eWxlPSJlbmFibGUtYmFja2dyb3VuZDpuZXcgMCAwIDEyOCAxMjg7IiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjAgMCAxMjggMTI4IiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIj48c3R5bGUgdHlwZT0idGV4dC9jc3MiPgoJLnN0MHtmaWxsOiMzMUFGOTE7fQoJLnN0MXtmaWxsOiNGRkZGRkY7fQo8L3N0eWxlPjxnPjxjaXJjbGUgY2xhc3M9InN0MCIgY3g9IjY0IiBjeT0iNjQiIHI9IjY0Ii8+PC9nPjxnPjxwYXRoIGNsYXNzPSJzdDEiIGQ9Ik01NC4zLDk3LjJMMjQuOCw2Ny43Yy0wLjQtMC40LTAuNC0xLDAtMS40bDguNS04LjVjMC40LTAuNCwxLTAuNCwxLjQsMEw1NSw3OC4xbDM4LjItMzguMiAgIGMwLjQtMC40LDEtMC40LDEuNCwwbDguNSw4LjVjMC40LDAuNCwwLjQsMSwwLDEuNEw1NS43LDk3LjJDNTUuMyw5Ny42LDU0LjcsOTcuNiw1NC4zLDk3LjJ6Ii8+PC9nPjwvc3ZnPg==",
      error_imgPath: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiA/PjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMC8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvVFIvMjAwMS9SRUMtU1ZHLTIwMDEwOTA0L0RURC9zdmcxMC5kdGQnPjxzdmcgaGVpZ2h0PSIzMiIgc3R5bGU9Im92ZXJmbG93OnZpc2libGU7ZW5hYmxlLWJhY2tncm91bmQ6bmV3IDAgMCAzMiAzMiIgdmlld0JveD0iMCAwIDMyIDMyIiB3aWR0aD0iMzIiIHhtbDpzcGFjZT0icHJlc2VydmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiPjxnPjxnIGlkPSJFcnJvcl8xXyI+PGcgaWQ9IkVycm9yIj48Y2lyY2xlIGN4PSIxNiIgY3k9IjE2IiBpZD0iQkciIHI9IjE2IiBzdHlsZT0iZmlsbDojRDcyODI4OyIvPjxwYXRoIGQ9Ik0xNC41LDI1aDN2LTNoLTNWMjV6IE0xNC41LDZ2MTNoM1Y2SDE0LjV6IiBpZD0iRXhjbGFtYXRvcnlfeDVGX1NpZ24iIHN0eWxlPSJmaWxsOiNFNkU2RTY7Ii8+PC9nPjwvZz48L2c+PC9zdmc+"
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(READ_VIEW_REGISTRY_ID, () => new ReadView());
    this.quickViewNavigator.register(CREATE_VIEW_REGISTRY_ID, () => new CreateView());
    this.quickViewNavigator.register(UPDATE_VIEW_REGISTRY_ID, () => new UpdateView());
    this.quickViewNavigator.register(MESSAGE_VIEW_REGISTRY_ID, () => new MessageView());
    this.quickViewNavigator.register(DIALOGUE_VIEW_REGISTRY_ID, () => new DialogueView());

    this.cardNavigator.register(MEDIUM_VIEW_REGISTRY_ID, () => new MediumCardView());
    this.quickViewNavigator.register(SHOWALLMEDIUM_VIEW_REGISTRY_ID, () => new ShowAllView());
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo.svg');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'PeopleDetails-property-pane'*/
      './PeopleDetailsPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.PeopleDetailsPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    // return this.cardSize === 'Large' ? CARD_VIEW_REGISTRY_ID : MEDIUM_VIEW_REGISTRY_ID;
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
