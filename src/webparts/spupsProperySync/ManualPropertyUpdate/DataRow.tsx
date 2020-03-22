import * as React from 'react';
import EditableCell from './EditableCell';

export interface IDataRowProps {
    item: any;
    onTableUpdate: () => void;
    onDelRow: (item: any) => void;
}

export default function DataRow(props: IDataRowProps) {
    function onDelEvent() {
        props.onDelRow(props.item)
    }

    return (
        <tr className="eachRow">
            <EditableCell onTableUpdate={props.onTableUpdate} cellData={{
                "type": "name",
                value: props.item.name,
                id: props.item.id
            }} />
            <EditableCell onTableUpdate={props.onTableUpdate} cellData={{
                type: "price",
                value: props.item.price,
                id: props.item.id
            }} />
            <EditableCell onTableUpdate={props.onTableUpdate} cellData={{
                type: "qty",
                value: props.item.qty,
                id: props.item.id
            }} />
            <EditableCell onTableUpdate={props.onTableUpdate} cellData={{
                type: "category",
                value: props.item.category,
                id: props.item.id
            }} />
            <td className="del-cell">
                <input type="button" onClick={onDelEvent} value="X" className="del-btn" />
            </td>
        </tr>
    );
}