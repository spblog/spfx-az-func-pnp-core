import * as React from 'react';
import styles from './CallPnPCoreApi.module.scss';
import { ICallPnPCoreApiProps } from './ICallPnPCoreApiProps';
import { FC } from 'react';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export const CallPnPCoreApi: FC<ICallPnPCoreApiProps> = (props) => {
  const [lists, setLists] = React.useState<any[]>();

  React.useEffect(() => {
    props.context.aadHttpClientFactory
      .getClient('ca226d3c-f06d-4ea5-8bb4-f7b9b11df7da')
      .then((client: AadHttpClient): void => {
        client
          .get(`https://45c4bdee5fe9.ngrok.io/api/GetLists/?siteUrl=${props.context.pageContext.site.absoluteUrl}&tenantId=${props.context.pageContext.aadInfo.tenantId}`, AadHttpClient.configurations.v1)
          .then((response: HttpClientResponse): Promise<any[]> => {
            return response.json();
          })
          .then((response: any[]): void => {
            setLists(response);
          });
      });
  }, []);

  if (!lists) {
    return (
      <div>Loading....</div>
    );
  }

  return (
    <div className={styles.callPnPCoreApi}>
      <div>Site lists:</div>
      <ul>
        {lists.map(l => (
          <li>{l.title}</li>
        ))}
      </ul>
    </div>
  );
};
