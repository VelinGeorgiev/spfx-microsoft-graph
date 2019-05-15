import * as React from 'react';
import styles from './AadHttpClient.module.scss';
import { IAadHttpClientProps } from './IAadHttpClientProps';
import { IAadHttpClientState } from './IAadHttpClientState';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export default class AadHttpClientComponent extends React.Component<IAadHttpClientProps, IAadHttpClientState> {
  constructor(props: IAadHttpClientProps) {
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

    // Using Graph here, but any 1st or 3rd party REST API that requires Azure AD auth can be used here.
    this.props.aadHttpClientFactory
      .getClient('https://graph.microsoft.com')
      .then((client: AadHttpClient): Promise<HttpClientResponse> => {
        // Search for the users with givenName, surname, or displayName equal to the searchFor value
        return client
          .get(
            `https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName`,
            AadHttpClient.configurations.v1
          );
          //&$filter=(givenName%20eq%20'${escape(this.state.searchFor)}')%20or%20(surname%20eq%20'${escape(this.state.searchFor)}')%20or%20(displayName%20eq%20'${escape(this.state.searchFor)}')
      })
      .then(response => {
        return response.json();
      })
      .then(json => {

        // Prepare the output array
        var users: Array<any> = new Array<any>();

        // Log the result in the console for testing purposes
        console.log(json);

        // Map the JSON response to the output array
        json.value.map((item: any) => {
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
      })
      .catch(error => {
        console.error(error);
      });
  }

  public render(): React.ReactElement<IAadHttpClientProps> {
    return (
      <div className={styles.aadHttpClient}>
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
