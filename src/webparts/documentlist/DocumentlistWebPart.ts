import * as React from "react"
import * as ReactDom from "react-dom"
import { Version } from "@microsoft/sp-core-library"
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane"
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base"
import { IReadonlyTheme } from "@microsoft/sp-component-base"

import * as strings from "DocumentlistWebPartStrings"
import Documentlist from "./components/Documentlist"
import { IDocumentlistProps } from "./components/IDocumentlistProps"

export interface IDocumentlistWebPartProps {
  description: string
}

export default class DocumentlistWebPart extends BaseClientSideWebPart<IDocumentlistWebPartProps> {
  private _isDarkTheme: boolean = false
  private _environmentMessage: string = ""

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage()

    return super.onInit()
  }

  public render(): void {
    const element: React.ReactElement<IDocumentlistProps> = React.createElement(
      Documentlist,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
      }
    )

    ReactDom.render(element, this.domElement)
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return
    }

    this._isDarkTheme = !!currentTheme.isInverted
    const { semanticColors } = currentTheme
    this.domElement.style.setProperty("--bodyText", semanticColors.bodyText)
    this.domElement.style.setProperty("--link", semanticColors.link)
    this.domElement.style.setProperty(
      "--linkHovered",
      semanticColors.linkHovered
    )
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement)
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    }
  }
}
