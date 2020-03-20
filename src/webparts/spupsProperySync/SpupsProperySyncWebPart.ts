import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import { sp } from '@pnp/sp';
import { graph } from "@pnp/graph";

import * as strings from 'SpupsProperySyncWebPartStrings';
import SpupsProperySync from './components/SpupsProperySync';
import { ISpupsProperySyncProps } from './components/ISpupsProperySyncProps';
import * as jQuery from 'jquery';

export interface ISpupsProperySyncWebPartProps {
    context: WebPartContext;
}

export default class SpupsProperySyncWebPart extends BaseClientSideWebPart<ISpupsProperySyncWebPartProps> {

    protected async onInit() {
        await super.onInit();
        sp.setup(this.context);
        graph.setup({ spfxContext: this.context });

        jQuery("#workbenchPageContent").prop("style", "max-width: none");
        jQuery(".SPCanvas-canvas").prop("style", "max-width: none");
        jQuery(".CanvasZone").prop("style", "max-width: none");
    }

    public render(): void {
        const element: React.ReactElement<ISpupsProperySyncProps> = React.createElement(
            SpupsProperySync,
            {
                context: this.context
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
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
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
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
