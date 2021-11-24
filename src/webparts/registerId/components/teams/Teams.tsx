import * as React from 'react';
import { ITeam } from '../../models/team';
import { IRegisterIdProps } from '../IRegisterIdProps';
import Member from '../member/Member';
import styles from './Teams.module.scss';

export interface TeamsProps extends IRegisterIdProps {
  teams: ITeam[];
}

const Teams: React.FunctionComponent<TeamsProps> = ({ teams, context, siteUrl, spHttpClient, description, listName }: TeamsProps) => (
  <div className= { styles.teams }>
    <h1>Registered teams</h1>
    <div className={ styles.table }>
      <div className={`${ styles.row } ${ styles.header }`}>
        <div className={ styles.col }>Team name</div>
        <div className={ styles.col }>Description</div>
        <div className={ styles.col }>Project Url</div>
        <div className={ styles.col }>Team members</div>
      </div>
      { teams.map((team, index) =>
        <div className={`${ styles.row } ${ index % 2 > 0 ? styles.odd : '' }`}>
          <div className={ styles.col }>{ team.Title }</div>
          <div className={ styles.col }>{ team.Description }</div>
          <div className={ styles.col }>
            <a href={team.Project_x0020_link?.Url}>{team.Project_x0020_link?.Description}</a>
          </div>
          <div className={ styles.col }>
            { team.MembersId.map((member) =>
              <Member
                Id={member}
                context={context}
                siteUrl={siteUrl}
                spHttpClient={spHttpClient}
                description={description}
                listName={listName}
              ></Member>
            )}
          </div>
        </div>
      )}
    </div>
  </div>
);

export default Teams;