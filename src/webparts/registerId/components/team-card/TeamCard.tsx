import * as React from 'react';
import { ITeam } from '../../models/team';
import { IRegisterIdProps } from '../IRegisterIdProps';
import MemberCard from '../member-card/MemberCard';
import styles from './TeamCard.module.scss';
import { DefaultButton, Toggle } from 'office-ui-fabric-react';
import { useState } from 'react';
import TeamForm from '../team-form/TeamForm';

export interface TeamCardProps extends ITeam, IRegisterIdProps {
  reloadTeams: () => void;
}
 
const TeamCard: React.FunctionComponent<TeamCardProps> = ({ Id, Description, MembersId, Title, Project_x0020_link, AppFw, context, siteUrl, spHttpClient, description, listName, reloadTeams }: TeamCardProps) => {
  const [editMode, setEditMode] = useState(false);

  const enableEdit = () => {
    setEditMode(true);
  };

  const reloadTeamsAndReset = () => {
    reloadTeams();
    setEditMode(false);
  };

  const cancelUpdate = () => {
    setEditMode(false);
  };

  return (
    <>
    { !editMode &&
      <div className={ styles.teamCard }>
        <h1>Your project is registered!</h1>
        <div className= { styles.flex }>
          <p className={ styles.title }>Team name</p>
          <p>{Title}</p>
        </div>
        <div className= { styles.flex }>
          <p className={ styles.title }>Project Url</p>
          <p><a target='_blank' href={Project_x0020_link?.Url}>{Project_x0020_link?.Description}</a></p>
        </div>
        <p className={ styles.title }>Description</p>
        <p>{Description}</p>
        <p className={ styles.title }>Team members</p>
        <div className={ styles.members }>
          { MembersId &&
            MembersId.map((member) =>
              <MemberCard
                Id={member}
                context={context}
                siteUrl={siteUrl}
                spHttpClient={spHttpClient}
                description={description}
                listName={listName}
              ></MemberCard>)
          }
        </div>
        <div className= { styles.flex }>
          <p className={ styles.title }>Dynatrace App?</p>
          <p className={ styles.paddingTop }><Toggle label="" onText="Yes" offText="No" disabled={true} defaultChecked={AppFw}/></p>
        </div>
        <div className= { styles.flex } style={{marginTop: '-5px'}}><i>Want to know more about <a href='https://dynatrace.sharepoint.com/sites/Inno_Days/SitePages/Platform-Apps.aspx'>&nbsp;Dynatrace Apps?</a></i></div>
        <DefaultButton className={ styles.edit } text="Edit your team" allowDisabledFocus onClick={enableEdit}/>
      </div>
    }
    { editMode &&
      <TeamForm
        Id={Id}
        Description={Description}
        MembersId={MembersId}
        Title={Title}
        Project_x0020_link={Project_x0020_link}
        AppFw={AppFw}
        context={context}
        siteUrl={siteUrl}
        spHttpClient={spHttpClient}
        description={description}
        listName={listName}
        reloadTeams={reloadTeamsAndReset}
        cancelUpdate={cancelUpdate}
      ></TeamForm>
    }
    </>
  );
};

export default TeamCard;
