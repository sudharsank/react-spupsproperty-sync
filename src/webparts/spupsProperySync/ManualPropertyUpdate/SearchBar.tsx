import * as React from 'react'

export interface ISearchBarProps {
    onUserInput: (filterText: string) => void;    
}

export interface ISearchBarState {
    filterText: string;
}

export default class SearchBar extends React.Component<ISearchBarProps, ISearchBarState> {
    constructor(props: ISearchBarProps) {
        super(props);
        this.state = {
            filterText: ""
        };
    }
    private handleChange = () => {
        this.props.onUserInput(this.state.filterText);
    }

    public render(): JSX.Element {
        return (
            <div>
                <input type="text" placeholder="Search..." value={this.state.filterText} onChange={this.handleChange} />
            </div>
        )
    }
}