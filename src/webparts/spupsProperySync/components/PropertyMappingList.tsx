import * as React from 'react';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import { List } from 'office-ui-fabric-react/lib/List';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { Icon, IIconStyles, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import styles from './SpupsProperySync.module.scss';
import { IPropertyMappings } from '../../../Common/IModel';
import SPHelper from '../../../Common/SPHelper';
import * as _ from 'lodash';
import { parse } from 'json2csv';

const iconStyles: IIconStyles = {
	root: {
		fontSize: '24px',
		height: '24px',
		width: '24px'
	}
};
const downloadIcon: IIconProps = { iconName: 'SaveTemplate' };
const csvIcon: IIconProps = { iconName: 'FileTemplate' };

export interface IPropertyMappingProps {
	mappingProperties: IPropertyMappings[];
	helper: SPHelper;
}

export interface IPropertyMappingState {
	isOpen: boolean;
	templateProperties: IPropertyMappings[];
	downloadLink: string;
	templateFileName: string;
}

export default class PropertyMappingList extends React.Component<IPropertyMappingProps, IPropertyMappingState> {
	/**
	 * Default constructor
	 * @param props Component props
	 */
	constructor(props: IPropertyMappingProps) {
		super(props);
		this.state = {
			isOpen: false,
			templateProperties: [],
			downloadLink: '',
			templateFileName: ''
		}
	}
	/**
	 * Component mount
	 */
	public componentDidMount = () => {
		this.setState({ templateProperties: this.getDefaultTemplateProperties() });
	}
	/**
	 * Component updated
	 */
	public componentDidUpdate = (prevProps: IPropertyMappingProps) => {
		if (prevProps.mappingProperties !== this.props.mappingProperties) {
			this.setState({ templateProperties: this.getDefaultTemplateProperties() });
		}
	}
	/**
	 * Get the property mappings from the props
	 */
	private getDefaultTemplateProperties = () => {
		return this.props.mappingProperties;
	}
	/**
	 * Update the property mappings state by enabling or disabling the property
	 * Based on this the templates will be generated
	 */
	private _onEnableOrDisableProperty = (item: IPropertyMappings, checked: boolean) => {
		let templateProperties: IPropertyMappings[] = this.state.templateProperties;
		let property = templateProperties.filter(prop => { return prop.ID == item.ID; });
		if (property) property[0].IsIncluded = false;
		this.setState({ templateProperties });
		//this.render();
		console.log(this.state.templateProperties);
	}
	/**
	 * Get the default property mappings and then open the panel
	 */
	private _openPropertyMappingPanel = () => {
		let templateProperties = this.getDefaultTemplateProperties();
		templateProperties.map(prop => { prop.IsIncluded = true; });
		this.setState({ templateProperties, isOpen: true });
	}
	/**
	 * Dismiss or close the panel
	 */
	private _dismissPanel = () => {
		this.setState({ isOpen: false });
	}
	/**
	 * Custom panel footer contents with buttons
	 */
	private _onRenderPanelFooterContent = (): JSX.Element => {
		return (
			<div className={styles.panelFooter}>
				<PrimaryButton iconProps={downloadIcon} onClick={this._generateJSONTemplate}>Generate JSON</PrimaryButton>
				<PrimaryButton iconProps={csvIcon} onClick={this._generateCSVTemplate}>Generate CSV</PrimaryButton>
			</div>
		);
	}
	private _getIncludedPropertyMapping = () => {
		return _.filter(this.state.templateProperties, (o) => { return o.IsIncluded; });
	}
	/**
	 * Button click to generate the JSON template
	 */
	private _generateJSONTemplate = async () => {
		const { helper } = this.props;
		let jsonOut = await helper.getPropertyMappingsTemplate(this._getIncludedPropertyMapping());
		let fileTemplate = await helper.addFilesToFolder(JSON.stringify(jsonOut));
		this.setState({
			downloadLink: fileTemplate.data.ServerRelativeUrl,
			templateFileName: fileTemplate.data.Name
		}, this.getTemplateFile);
	}
	/**
	 * Download the JSON file
	 */
	private getTemplateFile = async () => {
		let blobContent: any = await this.props.helper.getFileContentAsBlob(this.state.downloadLink);
		if (window.navigator.msSaveOrOpenBlob) {
			window.navigator.msSaveBlob(blobContent, this.state.templateFileName);
		} else {
			const anchor = window.document.createElement('a');
			anchor.href = window.URL.createObjectURL(blobContent);
			anchor.download = this.state.templateFileName;
			document.body.appendChild(anchor);
			anchor.click();
			document.body.removeChild(anchor);
		}
	}
	/**
	 * Button click to generate the CSV template
	 */
	private _generateCSVTemplate = async () => {
		const { helper } = this.props;
		let templateProperties = this._getIncludedPropertyMapping();
		let fields: string[] = [];
		fields.push("UserID");
		templateProperties.map(propmap => {
			fields.push(propmap.SPProperty);
		});
		const csv = parse("", { fields });
		let fileTemplate = await helper.addFilesToFolder(csv);
		this.setState({
			downloadLink: fileTemplate.data.ServerRelativeUrl,
			templateFileName: fileTemplate.data.Name
		}, this.getTemplateFile);
	}
	/**
	 * Render the property mapping item in the List
	 */
	private _onRenderCell = (item: IPropertyMappings, index: number | undefined): JSX.Element => {
		return (
			<div className={styles.mappingcontainer} data-is-focusable={true}>
				<div className={styles.propertydiv}>{item.AzProperty}</div>
				<Separator className={styles.separator}>
					<Icon iconName="DoubleChevronRight8" styles={iconStyles} />
				</Separator>
				<div className={styles.propertydiv}>{item.SPProperty}</div>
				<div className={styles.togglediv}>
					<Toggle label="" defaultChecked onChange={(e, checked) => { this._onEnableOrDisableProperty(item, checked); }} />
				</div>
			</div>
		);
	}
	/**
	 * Component render
	 */
	public render(): JSX.Element {
		const { isOpen, templateProperties } = this.state;
		return (
			<div className={styles.propertyMappingList}>
				<PrimaryButton text="Generate Template" onClick={this._openPropertyMappingPanel} />
				<Panel isOpen={isOpen} onDismiss={this._dismissPanel} type={PanelType.largeFixed} closeButtonAriaLabel="Close" headerText={"Property Mappings"}
					headerClassName={styles.panelHeader} isFooterAtBottom={true} onRenderFooterContent={this._onRenderPanelFooterContent}>
					<div className={styles.propertyMappingPanelContent}>
						<div className={styles.mappingcontainer} data-is-focusable={true} style={{ marginBottom: '10px' }}>
							<div className={styles.propertytitlediv}>{"Azure Property"}</div>
							<div className={styles.separator}>&nbsp;</div>
							<div className={styles.propertytitlediv}>{"SharePoint Property"}</div>
							<div className={styles.propertytitlediv} style={{ padding: '0px' }}>{"Enabled/Disabled"}</div>
						</div>
						<List items={templateProperties} onRenderCell={this._onRenderCell} />
					</div>
				</Panel>
			</div>
		)
	}
}