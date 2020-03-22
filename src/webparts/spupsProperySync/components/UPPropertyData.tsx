import * as React from 'react';
import styles from './SpupsProperySync.module.scss';
import * as strings from 'SpupsProperySyncWebPartStrings';
import { DetailsList, buildColumns, IColumn, DetailsListLayoutMode, ConstrainMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import MessageContainer from './MessageContainer';
import { MessageScope } from '../../../Common/IModel';
import * as csv from 'csvtojson';

const jsonData: any = `[
    {
        "UserID": "user1@tenantname.onmicrosoft.com",
        "Department": "department 1",
        "Office": "India",
        "CellPhone": "345345",
        "StreetAddress": "sssdfsdf s",
        "PostalCode": ""
    },
    {
        "UserID": "user2@tenantname.onmicrosoft.com",
        "Department": "",
        "Office": "Singapore",
        "CellPhone": "3434534",
        "StreetAddress": "f dfg dfg",
        "PostalCode": "345456"
    }
]`;
const csvData = `UserID,"Department, SPS-Department","Title, SPS-jobTitle","Office, SPS-Location",workPhone,CellPhone,Fax,StreetAddress,City,State,PostalCode,Country
user1@tenantname.onmicrosoft.com,dept1,Title 1,Office 1,324234234,455454,34234,Street address 1,City 1,State 1,876987,Country 1
user2@tenantname.onmicrosoft.com,dept2,Title 2,Office 2,234233423,343434,234234,street address 2,City 2,State 2,234567,Country 2`;

export interface IUPPropertyDataProps {
	items: any;
	isCSV: boolean;
}

export interface IUPPropertyDataState {
	items: any;
	columns: IColumn[];
	dynamicColumns: string[];
	searchText: string;
	emptyValues: boolean;
}

export default class UPPropertyData extends React.Component<IUPPropertyDataProps, IUPPropertyDataState> {
	private emptyValues: boolean = false;
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
		this.emptyValues = false;
		let cols: IColumn[] = [];
		if (columns && columns.length > 0) {
			columns.map((col: string) => {
				if (col.toLocaleLowerCase() == "userid") {
					cols.push({ key: col, name: col, fieldName: col, minWidth: 300, maxWidth: 300 } as IColumn);
				} else {
					cols.push({
						key: col, name: col, fieldName: col,
						onRender: (item: any, index: number, column: IColumn) => {
							if (item[col]) {
								return (<div>{item[col]}</div>);
							} else {
								this.emptyValues = true;
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
		// console.log(this.props.items);
		// console.log(this.props.isCSV);
		let finalOut = await csv().fromString(csvData);
		//console.log(finalOut);
		this._getJSONData(finalOut);
	}

	private _getJSONData = (inputjson?: any) => {
		let parsedJson = (inputjson) ? inputjson : JSON.parse(jsonData);
		let _dynamicColumns: string[] = [];
		Object.keys(parsedJson[0]).map((key) => {
			_dynamicColumns.push(key);
		});
		this.setState({
			columns: this._buildColumns(_dynamicColumns),
			items: parsedJson,
			emptyValues: this.emptyValues
		});
	}

	public render(): JSX.Element {
		const { items, columns, emptyValues } = this.state;
		//console.log(this.emptyValues);
		return (
			<div className={styles.uppropertydata}>
				{this.emptyValues &&
					<MessageContainer MessageScope={MessageScope.Info} Message={strings.EmptyDataWarningMsg} />
				}
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
					className={styles.uppropertylist}/>
			</div>
		);
	}
}