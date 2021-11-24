import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { sp } from '@pnp/sp';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from 'react';
import { useEffect, useState } from 'react';
import { MemberCardProps } from '../member-card/MemberCard';

const Member: React.FunctionComponent<MemberCardProps> = ({ Id, context, spHttpClient, siteUrl }: MemberCardProps) => {
  const [userId, _] = useState(Id);
  const [user, setUser] = useState<{Email: string; Title: string;}>();

  useEffect(() => {
    sp.setup({
      spfxContext: context
    });
    getUser();
  }, []);

  const getUser = (): void => {
    spHttpClient.get(`${siteUrl}/_api/web/getuserbyid('${userId}')`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
      },
    })
    .then((user: SPHttpClientResponse): Promise<any> => user.json())  
    .then((user: any): void => {
      setUser(user);
    })
    .catch((error: any): void => setUser(undefined));
  };

  return (
    <div>
      <Persona
        text={user?.Title}
        size={PersonaSize.size8}
        hidePersonaDetails={true}
        imageAlt="No presence detected"
      />
    </div>
  )
}

export default Member;
