import * as React from 'react';
import styles from './RegisterId.module.scss';
import { IRegisterIdProps } from './IRegisterIdProps';
import { sp } from '@pnp/sp';
import { useState, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ISPListTeams, ITeam } from '../models/team';
import TeamCard from './team-card/TeamCard';
import Teams from './teams/Teams';
import TeamForm from './team-form/TeamForm';

const descending = (a, b)=> {
  if ( a.Modified > b.Modified ){
    return -1;
  }
  if ( a.Modified < b.Modified ){
    return 1;
  }
  return 0;
}
 
const RegisterId: React.FunctionComponent<IRegisterIdProps> = ({ description, context, spHttpClient, siteUrl, listName }: IRegisterIdProps) => {
  const [teams, setTeams] = useState([])
  const [userId, setUserId] = useState()
  const [myTeam, setMyTeam] = useState<ITeam>()
  const [loading, setLoading] = useState(true);
  const [loadingTeam, setLoadingTeam] = useState(true);
  
  useEffect(() => {
    sp.setup({
      spfxContext: context
    });
    getTeams();
  }, []);

  useEffect(() => {
    getUserId();
  }, [teams]);

  useEffect(() => {
    if (!loading) {
      setTeam();
    }
  }, [loading]);

  const getUserId = (): void => {
    const body: string = JSON.stringify({
      'logonName': context.pageContext.user.loginName,     
    });

    spHttpClient.post(`${siteUrl}/_api/web/ensureuser`, SPHttpClient.configurations.v1, { body })
    .then((response: SPHttpClientResponse) => response.json())
    .then((user) => {
      setUserId(user.Id)
      setLoading(false);
    });
  };
  
  const getTeams = (): void => {
    setLoading(true);
    setLoading(true);
    setTeams([]);
    spHttpClient.get(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`,  
    SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },
    })  
    .then((response: SPHttpClientResponse): Promise<ISPListTeams> => response.json())
    .then((response) => { 
      return response;
    }) 
    .then((teams: ISPListTeams): void => setTeams(
      teams.value.sort(descending).map((team) => ({
        Title: team.Title,
        Id: team.Id,
        MembersId: team.MembersId,
        Description: team.Description,
        Project_x0020_link: team.Project_x0020_link,
      }))
    ))
    .catch((error: any): void => setTeams([])); 
  }

  const setTeam = () => {
    const currentTeam = teams?.filter((team) => team.MembersId?.indexOf(userId) > -1);

    setMyTeam(currentTeam.length > 0 ? {...currentTeam[0]} : undefined);
    setLoadingTeam(false);
  }

  return (
    <>
      <div className={ styles.registerId }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              { loadingTeam && <div className={ styles.loading }>Getting Team</div>}
              { !loadingTeam &&
                <>
                  { myTeam &&
                    <TeamCard
                      Id={myTeam?.Id}
                      Title={myTeam?.Title}
                      Description={myTeam?.Description}
                      MembersId={myTeam?.MembersId}
                      Project_x0020_link={myTeam?.Project_x0020_link}
                      context={context}
                      siteUrl={siteUrl}
                      spHttpClient={spHttpClient}
                      description={description}
                      listName={listName}
                      reloadTeams={getTeams}
                    ></TeamCard>
                  }
                  { !myTeam &&
                    <>
                      <TeamForm
                        MembersId={[userId]}
                        context={context}
                        siteUrl={siteUrl}
                        spHttpClient={spHttpClient}
                        description={description}
                        listName={listName}
                        reloadTeams={getTeams}
                      ></TeamForm>
                    </>
                  }
                </>
              }
            </div>
          </div>
        </div>
      </div>
      { loading && <div className={ styles.loading }>Getting teams</div>}
      { !loading &&
        <Teams
          teams={teams}
          description={description}
          context={context}
          spHttpClient={spHttpClient}
          siteUrl={siteUrl}
          listName={listName}
        ></Teams>
      }
    </>
  )
}

export default RegisterId;
