import * as React from 'react';
import styles from './PnPJsMsGraph.module.scss';
import { IPnPJsMsGraphProps } from './IPnPJsMsGraphProps';
import { graph } from "@pnp/graph";
import { IPnPJsMsGraphState } from './IPnPJsMsGraphState';

export default class PnPJsMsGraph extends React.Component<IPnPJsMsGraphProps, IPnPJsMsGraphState> {

  constructor(props: IPnPJsMsGraphProps) {
    super(props);
    this.state = { listItems: [] };
  }

  public componentDidMount(): void {
    graph.users
      .select("displayName,mail,userPrincipalName")
      .get()
      .then(users => {
        this.setState({ listItems: users });
      });
  }

  public render(): React.ReactElement<IPnPJsMsGraphProps> {
    return (
      <div className={styles.pnPJsMsGraph}>
        <div className={styles.container}>
          <div className={styles.row}>
            <span className={styles.title}>Welcome to the PnP Possum Pete team!</span>
            {
              this.state.listItems.map(x => {
                return <div>{JSON.stringify(x)}</div>;
              })
            }
          </div>
        </div>
      </div>
    );
  }
}
