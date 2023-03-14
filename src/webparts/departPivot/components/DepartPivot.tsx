import * as React from 'react';
import styles from './DepartPivot.module.scss';
import { IDepartPivotProps } from './IDepartPivotProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";
import NewPivot from './NewPivot';
import { MSGraphClient } from "@microsoft/sp-http";
export default class DepartPivot extends React.Component<IDepartPivotProps, {}> {
  constructor(prop: IDepartPivotProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
    graph.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<IDepartPivotProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
    <NewPivot context={this.props.context} propertyToggle={this.props.propertyToggle}/>
    );
  }
}
