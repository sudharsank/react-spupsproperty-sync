import * as React from 'react';
import EditableCell from './EditableCell';

export interface IDataRowProps {
    item: any;
    columns: any;
    onTableUpdate: () => void;
    onDelRow: (item: any) => void;
}

export default function DataRow(props: IDataRowProps) {
    function onDelEvent() {
        props.onDelRow(props.item);
    }
    return (
        <tr className="eachRow">
            {props.columns.map(col => {
                if(col !== "ImageUrl") {
                    if (col == "UserID") {
                        return <EditableCell onTableUpdate={props.onTableUpdate} cellData={{
                            "type": col,
                            value: props.item.DisplayName,
                            id: props.item.UserID,
                            label: true,
                            ImageUrl: props.item.ImageUrl
                        }} />;
                    } else {
                        return <EditableCell onTableUpdate={props.onTableUpdate} cellData={{
                            "type": col,
                            value: props.item[col],
                            id: props.item.UserID,
                            label: false
                        }} />;
                    }
                }                
            })}
            <td className="del-cell">
                <input type="button" onClick={onDelEvent} value="X" className="del-btn" />
            </td>
        </tr>
    );
}