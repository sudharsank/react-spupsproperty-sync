import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { sp } from '@pnp/sp';
import { graph } from "@pnp/graph";
import * as strings from 'SpupsProperySyncWebPartStrings';
import SpupsProperySync from './components/SpupsProperySync';
import { ISpupsProperySyncProps } from './components/SpupsProperySync';

export interface ISpupsProperySyncWebPartProps {
    context: WebPartContext;
    templateLib: string;
    appTitle: string;
    AzFuncUrl: string;
    UseCert: boolean;
}

export default class SpupsProperySyncWebPart extends BaseClientSideWebPart<ISpupsProperySyncWebPartProps> {

    protected async onInit() {
        await super.onInit();
        sp.setup(this.context);
        graph.setup({ spfxContext: this.context });

        // jQuery("#workbenchPageContent").prop("style", "max-width: none");
        // jQuery(".SPCanvas-canvas").prop("style", "max-width: none");
        // jQuery(".CanvasZone").prop("style", "max-width: none");
    }

    public render(): void {
        const element: React.ReactElement<ISpupsProperySyncProps> = React.createElement(
            SpupsProperySync,
            {
                context: this.context,
                templateLib: this.properties.templateLib,
                displayMode: this.displayMode,
                appTitle: this.properties.appTitle,
                AzFuncUrl: this.properties.AzFuncUrl,
                UseCert: this.properties.UseCert,
                updateProperty: (value: string) => {
                    this.properties.appTitle = value;
                },
                openPropertyPane: this.openPropertyPane
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected get disableReactivePropertyChanges() {
        return true;
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    private openPropertyPane = (): void => {
        this.context.propertyPane.open();
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
                                PropertyFieldListPicker('templateLib', {
                                    key: 'templateLibFieldId',
                                    label: strings.PropTemplateLibLabel,
                                    selectedList: this.properties.templateLib,
                                    includeHidden: false,
                                    orderBy: PropertyFieldListPickerOrderBy.Title,
                                    disabled: false,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    context: this.context,
                                    onGetErrorMessage: null,
                                    deferredValidationTime: 0,
                                    baseTemplate: 101,
                                    listsToExclude: ['Documents']
                                }),
                                PropertyPaneTextField('AzFuncUrl', {
                                    label: strings.PropAzFuncLabel,
                                    description: strings.PropAzFuncDesc,
                                    multiline: true,
                                    placeholder: strings.PropAzFuncLabel,
                                    resizable: true,
                                    rows: 5,
                                    value: this.properties.AzFuncUrl
                                }),
                                PropertyFieldToggleWithCallout('UseCert', {
                                    calloutTrigger: CalloutTriggers.Hover,
                                    key: 'UseCertFieldId',
                                    label: strings.PropUseCertLabel,
                                    calloutContent: React.createElement('div', {}, strings.PropUseCertCallout),
                                    onText: 'ON',
                                    offText: 'OFF',
                                    checked: this.properties.UseCert
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
