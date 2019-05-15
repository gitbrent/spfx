import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";

import * as strings from "SpFxAppWebPartStrings";
import SpFxApp from "./components/SpFxApp";
import { ISpFxAppProps } from "./components/ISpFxAppProps";

require("@pnp/logging");
require("@pnp/common");
require("@pnp/odata");
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";
import { sp } from "@pnp/sp";

export interface ISpFxAppWebPartProps {
  description: string;
}

export default class SpFxAppWebPart extends BaseClientSideWebPart<
  ISpFxAppWebPartProps
> {
	/**
     * Initialize the web part.
     */
    protected onInit(): Promise<void> {
      sp.setup({
        spfxContext: this.context
      });

      // optional, we are setting up the @pnp/logging for debugging
      Logger.activeLogLevel = LogLevel.Info;
      Logger.subscribe(new ConsoleListener());

      return super.onInit();
    }

  public render(): void {
    const element: React.ReactElement<ISpFxAppProps> = React.createElement(
      SpFxApp,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
