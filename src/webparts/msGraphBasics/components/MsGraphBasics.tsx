import * as React from 'react';
import styles from './MsGraphBasics.module.scss';
import { IMsGraphBasicsProps } from './IMsGraphBasicsProps';
import { IMsGraphBasicsState } from './IMsGraphBasicsState';
import { MSGraphClient } from '@microsoft/sp-http';

export default class MsGraphBasics extends React.Component<IMsGraphBasicsProps, IMsGraphBasicsState> {

  constructor(props: IMsGraphBasicsProps) {
    super(props);
    this.state = { listItems: [] };
  }

  public componentDidMount(): void {

    // this.props.aadHttpClientFactory
    // .getClient('53dabdc1-247c-43c4-86fc-5a005b850df5')
    // .then((client: AadHttpClient): Promise<HttpClientResponse> => {

    //   return client.get('https://localhost:44300/api/flavours', AadHttpClient.configurations.v1);
    // })
    // .then((response: HttpClientResponse): Promise<IceCream[]> => {

    //   return response.json();
    // })
    // .then((jsonResult: IceCream[]): void => {

    //   this.setState({ iceCreamList: jsonResult });
    // })
    // .catch(err => console.log(err));

    this.props.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api("users")
          .version("v1.0")
          .select("displayName,mail,userPrincipalName")
          //.filter(`(givenName eq '${escape(this.state.searchFor)}') or (surname eq '${escape(this.state.searchFor)}') or (displayName eq '${escape(this.state.searchFor)}')`)
          .get((err, res) => {  

            if (err) {
              console.error(err);
              return;
            }

            // Prepare the output array
            var users: Array<any> = new Array<any>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              users.push( { 
                displayName: item.displayName,
                mail: item.mail,
                userPrincipalName: item.userPrincipalName,
              });
            });

            // Update the component state accordingly to the result
            this.setState(
              {
                listItems: users,
              }
            );
          });
      });
  }
  
  public render(): React.ReactElement<IMsGraphBasicsProps> {
    return (
      <div className={styles.msGraphBasics}>
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
