import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldLabelWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldLabelWithCallout';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
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
    dateFormat: string;
    toggleInfoHeaderValue: boolean;
    useFullWidth: boolean;
}

export default class SpupsProperySyncWebPart extends BaseClientSideWebPart<ISpupsProperySyncWebPartProps> {

    protected async onInit() {
        await super.onInit();
        sp.setup(this.context);
        graph.setup({ spfxContext: this.context });        
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
                dateFormat: this.properties.dateFormat ? this.properties.dateFormat : "DD, MMM YYYY hh:mm A",
                useFullWidth: this.properties.useFullWidth,
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
                                PropertyPaneWebPartInformation({
                                    description: `${strings.PropInfoTemplateLib}`,
                                    key: 'templateLibInfoId'
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
                                }),
                                PropertyPaneWebPartInformation({
                                    description: `${strings.PropInfoUseCert}`,
                                    key: 'useCertInfoId'
                                }),
                                PropertyPaneTextField('dateFormat', {
                                    label: strings.PropDateFormatLabel,
                                    description: '',
                                    multiline: false,
                                    placeholder: strings.PropDateFormatLabel,
                                    resizable: false,
                                    value: this.properties.dateFormat
                                }),
                                PropertyPaneWebPartInformation({
                                    description: `${strings.PropInfoDateFormat}`,
                                    key: 'dateFormatInfoId'
                                }),
                                PropertyFieldToggleWithCallout('useFullWidth', {
                                    //calloutTrigger: CalloutTriggers.Hover,
                                    key: 'useFullWidthFieldId',
                                    label: 'Use page full width',
                                    //calloutContent: React.createElement('div', {}, strings.PropUseCertCallout),
                                    onText: 'ON',
                                    offText: 'OFF',
                                    checked: this.properties.useFullWidth
                                }),
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
