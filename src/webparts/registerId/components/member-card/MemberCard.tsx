import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { sp } from '@pnp/sp';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from 'react';
import { useEffect, useState } from 'react';
import { IMember } from '../../models/member';
import { IRegisterIdProps } from '../IRegisterIdProps';
import styles from './MemberCard.module.scss';

export interface MemberCardProps extends IMember, IRegisterIdProps {}

const personaStyles = {
  primaryText: { fontSize: '14px', color: 'white' },
  root: { margin: '10px' },
};
 
const MemberCard: React.FunctionComponent<MemberCardProps> = ({ Id, context, spHttpClient, siteUrl }: MemberCardProps) => {
  const [userId, _] = useState(Id);
  const [user, setUser] = useState<{Email: string; Title: string;}>();

  useEffect(() => {
    sp.setup({
      spfxContext: context
    });
    getUser();
  }, [userId]);

  const getUser = (): void => {
    spHttpClient.get(`${siteUrl}/_api/web/getuserbyid('${userId}')`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },
    })  
    .then((response: SPHttpClientResponse): Promise<any> => response.json())  
    .then((userResponse: any): void => {
      setUser(userResponse);
    })
    .catch((error: any): void => setUser(undefined)); 
  };
  
  return (
    <div className={ styles.memberCard }>
      <Persona
        imageUrl={`${siteUrl}/_layouts/15/userphoto.aspx?size=L&username=${user?.Email}`}
        text={user?.Title}
        size={PersonaSize.size32}
        imageAlt={user?.Title}
        className={ styles.member }
        styles={personaStyles}
      />
    </div>
  );
};

export default MemberCard;
