import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneHorizontalRule,
    PropertyPaneTextField,
    PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import bfb from 'bfb';

import * as strings from 'BingSearchBoxWebPartStrings';

export interface IBingSearchBoxWebPartProps {
    width: number;
    height: number;
    title: string;
    cornerRadius: number;
    strokeOutline: boolean;
    dropShadow: boolean;
    iconColor: string;
    ghostText: string;
}

export default class BingSearchBoxWebPart extends BaseClientSideWebPart<IBingSearchBoxWebPartProps> {

    private _isDarkTheme: boolean = false;
    private _environmentMessage: string = '';
    private _iconColor: string = "#067FA6";

    public render(): void {
        this.domElement.innerHTML = `<div id="bfb_searchbox"></div>`;

        const bfbSearchBoxConfig = {
            containerSelector: "bfb_searchbox",
            width: this.properties.width,                             // default: 560, min: 360, max: 650
            height: this.properties.height,                             // default: 40, min: 40, max: 72
            title: this.properties.title,               // default: "Bing search box"
            cornerRadius: this.properties.cornerRadius,                        // default: 6, min: 0, max: 25                                  
            strokeOutline: this.properties.strokeOutline,                    // default: true
            dropShadow: this.properties.dropShadow,                       // default: true
            iconColor: this._iconColor,                   // default: #067FA6
            companyNameInGhostText: this.properties.ghostText       // default: not specified
            // when absent, ghost text will be "Search work and the web"
            // when specified, text will be "Search the web and [Contoso]"
        };
        bfb.Embedded.SearchBox.init(bfbSearchBoxConfig);
    }

    protected onInit(): Promise<void> {
        this._environmentMessage = this._getEnvironmentMessage();

        return super.onInit();
    }



    private _getEnvironmentMessage(): string {
        if (!!this.context.sdks.microsoftTeams) { // running in Teams
            return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
        }

        return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
    }

    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
        if (!currentTheme) {
            return;
        }
        this._iconColor = currentTheme?.palette?.themeDark;

        this._isDarkTheme = !!currentTheme.isInverted;
        const {
            semanticColors
        } = currentTheme;

        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }

    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
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
                            //groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('ghostText', { label: "Search box text" }),
                                PropertyPaneHorizontalRule(),
                                PropertyFieldNumber('width', { key:"width", label: "Width", minValue: 360, maxValue: 650, value:this.properties.width }),
                                PropertyFieldNumber('height', { key:"height", label: "Height", minValue: 40, maxValue: 72, value:this.properties.height }),
                                PropertyPaneTextField('title', { label: "Accessibility title" }),
                                PropertyFieldNumber('cornerRadius', { key:"cornerRadius", label: "Corner radius", minValue: 0, maxValue: 25, value:this.properties.cornerRadius }),
                                PropertyPaneToggle('strokeOutline', { label: "Stroke outline" }),
                                PropertyPaneToggle('dropShadow', { label: "Drop shadow" }),                                
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
