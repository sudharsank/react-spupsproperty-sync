import * as React from 'react';
import styles from './DynamicTable.module.scss';
import DataRow from './DataRow';
import { IPropertyMappings } from '../../../../Common/IModel';

const map: any = require('lodash/map');
const union: any = require('lodash/union');

export interface IEditableTableProps {
    onTableUpdate?: () => void;
    onRowDel: () => void;
    filterText?: string;
    data: any[];
    isReadOnly?: boolean;
    propertyMappings?: IPropertyMappings[];
}

export default function EditableTable(props: IEditableTableProps) {

    function getColumns() {
        //var keyData = JSON.parse(JSON.stringify(props.data));
        let includedProps: IPropertyMappings[] = props.propertyMappings.filter((o) => { return o.IsIncluded; });
        let finalProps: any[] = map(includedProps, (o) => {
            return {
                'Title': o.Title,
                'Props': o.AzProperty
            };
        });
        let cols = union([{ 'Title': 'ID', 'Props': 'id' }, { 'Title': 'Display Name', 'Props': 'displayName' }, { 'Title': 'UPN', 'Props': 'userPrincipalName' }], finalProps);
        return cols;
    }

    var columns = getColumns();
    var rowitem: any = props.data.map((item) => {
        return (<DataRow item={item} columns={columns} onTableUpdate={props.onTableUpdate} onDelRow={props.onRowDel} key={item.UserID} isReadOnly={props.isReadOnly} />);
    });
    return (
        <div className={styles.dynamicTable}>
            <table className={styles.table}>
                <thead>
                    <tr>
                        {props.isReadOnly ? (
                            <>
                                {columns.map(key => {
                                    if (key.Props.toLocaleLowerCase() !== "imageurl" && key.Props.toLocaleLowerCase() !== "userprincipalname" && key.Props.toLocaleLowerCase() !== "id") {
                                        return (<th>{key.Title}</th>);
                                    }
                                })}
                            </>
                        ) : (
                                <>
                                    {columns.map(key => {
                                        if (key.Props.toLocaleLowerCase() !== "imageurl" && key.Props.toLocaleLowerCase() !== "displayname") {
                                            return (<th>{key.Title}</th>);
                                        }
                                    })}
                                </>
                            )}

                    </tr>
                </thead>
                <tbody>
                    {rowitem}
                </tbody>
            </table>
        </div>
    );
}