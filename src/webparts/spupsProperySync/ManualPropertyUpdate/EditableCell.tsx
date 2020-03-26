import * as React from 'react';
import styles from './DynamicTable.module.scss';

export interface IEditableCellProps {
    cellData: any;
    onTableUpdate: (event: any) => void;
}

export interface IEditableCellState {
    inputtext: string;
}

export default class EditableCell extends React.Component<IEditableCellProps, IEditableCellState> {
    constructor(props: IEditableCellProps) {
        super(props);
        this.state = {
            inputtext: ""
        }
    }

    // componentDidMount = () => {
    //     this.setState({ inputtext: this.props.cellData.value });
    // }

    // componentDidUpdate = (prevProps: IEditableCellProps) => {
    //     if (prevProps.cellData !== this.props.cellData) {
    //         this.setState({ inputtext: this.props.cellData.value });
    //     }
    // }

    private handleTextChange = (e) => {
        this.setState({ inputtext: e.target.value });
        this.props.onTableUpdate(e);
    }

    public render(): JSX.Element {
        const { cellData, onTableUpdate } = this.props;
        return (
            <td>
                {!cellData.label ? (
                    <input type='text' className={styles.textInput} name={cellData.type} id={cellData.id} value={this.state.inputtext} onChange={this.handleTextChange} />
                ) : (
                        <div className={styles.divusername}>
                            <img src={cellData.ImageUrl} />
                            <label>{cellData.value}</label>
                            {/* <span>{cellData.id.split('|')[2]}</span> */}
                        </div>
                    )}
            </td>
        );
    }

}