import * as React from 'react';
import { IRegisterIdProps } from '../IRegisterIdProps';
import styles from './TeamForm.module.scss';
import { PrimaryButton, TextField } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useState, useEffect } from 'react';
import { sp } from '@pnp/sp';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ITeam } from '../../models/team';

export interface TeamFormProps extends IRegisterIdProps {
  Title?: string;
	Id?: number;
	Description?: string;
	MembersId?: number[];
	Project_x0020_link?: { Description: string; Url: string };
  reloadTeams: () => void;
}
 
const TeamForm: React.FunctionComponent<TeamFormProps> =
  ({ Id, Description, MembersId, Title, Project_x0020_link, context, siteUrl, spHttpClient, listName, reloadTeams }: TeamFormProps) => {
    const [members, setMembers] = useState(MembersId ?? []);
    const [name, setName] = useState(Title);
    const [url, setUrl] = useState<{ Description: string; Url: string}>(Project_x0020_link);
    const [desc, setDesc] = useState(Description);
    const [disabled, setDisabled] = useState(true);
    const [defaultMembers, setDefaultMembers] = useState([]);
    const [invalidUrl, setInvalidUrl] = useState(undefined);

    useEffect(() => {
      sp.setup({
        spfxContext: context
      });
      if (MembersId?.length > 0) {
        getDefaultMembers();
      }
    }, []);

    useEffect(() => {
      const isNameFilled = (name ?? '').trim() !== '';
      const isUrlFilled = (url?.Url ?? '').trim() !== '';
      const isDescFilled = (desc ?? '').trim() !== '';
      const isMembersFilled = members.length > 0;
      const validUrl = invalidUrl === undefined;

      setDisabled(!(isNameFilled && isUrlFilled && isDescFilled && isMembersFilled && validUrl));
    }, [members, name, url, desc, invalidUrl]);

    const createTeam = () => {
      const body: string = JSON.stringify({  
        'Title': name,
        'MembersId': members,
        'Description': desc,
        'Project_x0020_link': url,
      }); 

      spHttpClient.post(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`,  
      SPHttpClient.configurations.v1,  
      {  
        headers: {  
          'Accept': 'application/json;odata=nometadata',  
          'Content-type': 'application/json;odata=nometadata',  
          'odata-version': ''  
        },  
        body: body  
      })  
      .then((response: SPHttpClientResponse): Promise<ITeam> => {  
        return response.json();  
      })  
      .then((item: ITeam): void => {   
        reloadTeams();
        setName(undefined);
        setUrl(undefined);
        setDesc(undefined);
        setMembers([]);
      }, (error: any): void => { });  
    }

    const updateTeam = () => {
      const body: string = JSON.stringify({  
        'Title': name,
        'MembersId': members,
        'Description': desc,
        'Project_x0020_link': url,
      }); 

      spHttpClient.post(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${Id})`,  
      SPHttpClient.configurations.v1,  
      {  
        headers: {  
          'Accept': 'application/json;odata=nometadata',  
          'Content-type': 'application/json;odata=nometadata',  
          'odata-version': '',  
          'IF-MATCH': '*',  
          'X-HTTP-Method': 'MERGE' 
        },  
        body: body  
      })
      .then((response: SPHttpClientResponse): void => reloadTeams())
      .catch((error: any): void => { }); 
    }

    const getDefaultMembers = () => {
      const filter = MembersId.map((member) => `(Id eq ${member})`).join(' or ');
      spHttpClient.get(`${siteUrl}/_api/web/lists/getbytitle('User Information List')/items?$filter=${filter}`,
      SPHttpClient.configurations.v1,  
      {  
        headers: {  
          'Accept': 'application/json;odata=nometadata',  
          'Content-type': 'application/json;odata=nometadata',  
          'odata-version': ''  
        },
      })  
      .then((response: SPHttpClientResponse): Promise<any> => response.json())
      .then((response) => {
        const emails = response.value.map((user) => user.EMail);
        setDefaultMembers(emails);
      })
      .catch((error: any): void => setDefaultMembers([]));
    }

    const validateUrl = (str): boolean => {
      var pattern = new RegExp('^(https?:\\/\\/)?'+ // protocol
        '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|'+ // domain name
        '((\\d{1,3}\\.){3}\\d{1,3}))'+ // OR ip (v4) address
        '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*'+ // port and path
        '(\\?[;&a-z\\d%_.~+=-]*)?'+ // query string
        '(\\#[-a-z\\d_]*)?$','i'); // fragment locator
      return !!pattern.test(str);
    }

    const getPeoplePickerItems = (items: any[]) => {
      setMembers(items.map((item) => item.id));
    }

    const setProjectName = (e) => {
      setName(e.target?.value)
    }

    const setProjectDescription = (e) => {
      setDesc(e.target?.value)
    }

    const setProjectUrl = (e) => {
      const url = e.target.value;
      if (validateUrl(url)) {
        setInvalidUrl(undefined);
        setUrl({ Description: url, Url: url})
      } else {
        setUrl({ Description: url, Url: url})
        setInvalidUrl('Provide a valid URL');
      }
    }

    return (
      <>
        <div className={ styles.teamForm }>
          { !Id &&
            <h1>Do you Have an idea?</h1>
          }
          <div className= { styles.flex }>
            <p className={ styles.title }>Team name</p>
            <p><TextField placeholder="What is the name of your team?" required value={name} onChange={setProjectName}/></p>
          </div>
          <div className= { styles.flex }>
            <p className={ styles.title }>Project Url</p>
            <p><TextField placeholder="What is the url of your project?" required value={url?.Url} onChange={setProjectUrl} errorMessage={invalidUrl}/></p>
          </div>
          <p className={ styles.title }>Description</p>
          <p><TextField multiline rows={3} required value={desc} onChange={setProjectDescription}/></p>
          <p className={ styles.title }>Team members</p>
          <p className={ styles.picker }>
            <PeoplePicker
              context={context}
              titleText="List the members of your team"
              personSelectionLimit={10}
              showtooltip={true}
              required={true}
              disabled={false}
              onChange={getPeoplePickerItems}
              showHiddenInUI={false}
              ensureUser={true}
              defaultSelectedUsers={defaultMembers}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000} />
          </p>
          { !Id &&
            <PrimaryButton
              text="Register team"
              className={ styles.save }
              allowDisabledFocus
              style={{backgroundColor:'#0078d4', border: 'none'}}
              onClick={createTeam}
              disabled={disabled}
            />
          }
          { Id &&
            <PrimaryButton
              text="Update team"
              className={ styles.save }
              allowDisabledFocus
              style={{backgroundColor:'#0078d4', border: 'none'}}
              onClick={updateTeam}
              disabled={disabled}
            />
          }
        </div>
      </>
    )
  };

export default TeamForm;
