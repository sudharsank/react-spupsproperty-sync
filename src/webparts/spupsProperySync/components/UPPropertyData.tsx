import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import * as strings from 'SpupsProperySyncWebPartStrings';
import { DetailsList, buildColumns, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import MessageContainer from './MessageContainer';
import { MessageScope } from '../../../Common/IModel';
import * as csv from 'csvtojson';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

const jsonData: any = `[
    {
        "UserID": "AdeleV@o365practice.onmicrosoft.com",
        "Department": "department 1",
        "Office": "India",
        "CellPhone": "345345",
        "StreetAddress": "sssdfsdf s",
        "PostalCode": ""
    },
    {
        "UserID": "GradyA@o365practice.onmicrosoft.com",
        "Department": "",
        "Office": "Singapore",
        "CellPhone": "3434534",
        "StreetAddress": "f dfg dfg",
        "PostalCode": "345456"
    }
]`;
const csvData = `UserID,"Department, SPS-Department","Title, SPS-jobTitle","Office, SPS-Location",workPhone,CellPhone,Fax,StreetAddress,City,State,PostalCode,Country
AdeleV@o365practice.onmicrosoft.com,dept1,Title 1,Office 1,324234234,455454,34234,Street address 1,City 1,State 1,876987,Country 1
GradyA@o365practice.onmicrosoft.com,dept2,Title 2,Office 2,234233423,343434,234234,street address 2,City 2,State 2,234567,Country 2`;

export interface IUPPropertyDataProps {
	items: any;
	isCSV: boolean;
	UpdateSPForBulkUsers: (data: any[]) => void;
}

export interface IUPPropertyDataState {
	items: any;
	columns: IColumn[];
	dynamicColumns: string[];
	searchText: string;
	emptyValues: boolean;
}

export default class UPPropertyData extends React.Component<IUPPropertyDataProps, IUPPropertyDataState> {
	constructor(props: IUPPropertyDataProps) {
		super(props);
		this.state = {
			items: [],
			columns: [],
			searchText: '',
			dynamicColumns: [],
			emptyValues: false
		};
	}

	public componentDidMount = () => {
		this._buildUploadDataList();
	}

	public componentDidUpdate = (prevProps: IUPPropertyDataProps) => {
		if (prevProps.items !== this.props.items || prevProps.isCSV !== this.props.isCSV) {
			this._buildUploadDataList();
		}
	}

	private _buildColumns = (columns: string[]): IColumn[] => {
		this.setState({ emptyValues: false });
		let cols: IColumn[] = [];
		if (columns && columns.length > 0) {
			columns.map((col: string) => {
				if (col.toLocaleLowerCase() == "userid") {
					cols.push({ key: col, name: col, fieldName: col, minWidth: 300, maxWidth: 300 } as IColumn);
				} else {
					cols.push({
						key: col, name: col, fieldName: col, minWidth: 150,
						onRender: (item: any, index: number, column: IColumn) => {
							if (item[col]) {
								return (<div>{item[col]}</div>);
							} else {
								this.setState({ emptyValues: true });
								return (<div className={styles.emptyData}>{strings.EmptyDataText}</div>);
							}
						}
					} as IColumn);
				}
			});
		}
		return cols;
	}

	private _buildUploadDataList = async () => {
		const { items, isCSV } = this.props;
		if (items) {
			if (isCSV) {
				let finalOut: any = await csv().fromString(items);
				this._getJSONData(finalOut);
			}
			else this._getJSONData(items);
		}
	}

	private _getJSONData = (inputjson?: any) => {
		let parsedJson = (inputjson) ? inputjson : JSON.parse(inputjson);
		let _dynamicColumns: string[] = [];
		Object.keys(parsedJson[0]).map((key) => {
			_dynamicColumns.push(key);
		});
		this.setState({
			columns: this._buildColumns(_dynamicColumns),
			items: parsedJson
		});
	}

	private _updatePropsForBulkUsers = () => {
		this.props.UpdateSPForBulkUsers(this.state.items);
	}

	public render(): JSX.Element {
		const { items, columns, emptyValues } = this.state;
		//console.log(this.emptyValues);
		return (
			<div className={styles.uppropertydata}>
				{emptyValues &&
					<MessageContainer MessageScope={MessageScope.Info} Message={strings.EmptyDataWarningMsg} />
				}
				{items && items.length > 0 &&
					<>
						<DetailsList
							items={items}
							setKey="set"
							columns={columns}
							compact={true}
							layoutMode={DetailsListLayoutMode.justified}
							constrainMode={ConstrainMode.unconstrained}
							isHeaderVisible={true}
							selectionMode={SelectionMode.none}
							enableShimmer={true}
							className={styles.uppropertylist} />
						<div style={{ padding: "10px" }}>
							<PrimaryButton text={strings.BtnUpdateUserProps} onClick={this._updatePropsForBulkUsers} style={{ marginRight: '5px' }} />
						</div>
					</>
				}
			</div>
		);
	}
}