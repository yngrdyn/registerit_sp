import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { sp } from '@pnp/sp';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DefaultButton, Dialog, DialogFooter, DialogType, PrimaryButton, TextField, Toggle } from 'office-ui-fabric-react';
import * as React from 'react';
import { useEffect, useState } from 'react';
import { ITeam } from '../../models/team';
import { IRegisterIdProps } from '../IRegisterIdProps';
import styles from './TeamForm.module.scss';

export interface TeamFormProps extends IRegisterIdProps {
  Title?: string;
	Id?: number;
	Description?: string;
	MembersId?: number[];
	Project_x0020_link?: { Description: string; Url: string };
  AppFw?: boolean;
  reloadTeams: () => void;
  cancelUpdate?: () => void;
}
 
const TeamForm: React.FunctionComponent<TeamFormProps> =
  ({ Id, Description, MembersId, Title, Project_x0020_link, AppFw, context, siteUrl, spHttpClient, listName, reloadTeams, cancelUpdate }: TeamFormProps) => {
    const [members, setMembers] = useState(MembersId ?? []);
    const [name, setName] = useState(Title);
    const [url, setUrl] = useState<{ Description: string; Url: string}>(Project_x0020_link);
    const [desc, setDesc] = useState(Description);
    const [disabled, setDisabled] = useState(true);
    const [defaultMembers, setDefaultMembers] = useState([]);
    const [invalidUrl, setInvalidUrl] = useState(undefined);
    const [appFw, setAppFw] = useState(AppFw);
    const [showWarning, setShowWarning] = useState(false);
    const [showLoadingPeople, setShowLoadingPeople] = useState(true);

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
        'AppFw': appFw,
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
    };

    const updateTeam = () => {
      const body: string = JSON.stringify({  
        'Title': name,
        'MembersId': members,
        'Description': desc,
        'Project_x0020_link': url,
        'AppFw': appFw,
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
    };

    const getDefaultMembers = () => {
      setShowLoadingPeople(true);
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
        setTimeout(() => {
          setShowLoadingPeople(false);
        }, emails.length * 250);
      })
      .catch((error: any): void => setDefaultMembers([]));
    };

    const validateUrl = (str): boolean => {
      var pattern = new RegExp('^(https?:\\/\\/)?'+ // protocol
        '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|'+ // domain name
        '((\\d{1,3}\\.){3}\\d{1,3}))'+ // OR ip (v4) address
        '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*'+ // port and path
        '(\\?[;&a-z\\d%_.~+=-]*)?'+ // query string
        '(\\#[-a-z\\d_]*)?$','i'); // fragment locator
      return !!pattern.test(str);
    };

    const getPeoplePickerItems = (items: any[]) => {
      if (items.length > 6) {
        setShowWarning(true);
      }
      setMembers(items.map((item) => item.id));
    };

    const setProjectName = (e) => {
      setName(e.target?.value);
    };

    const setProjectDescription = (e) => {
      setDesc(e.target?.value);
    };

    const setProjectUrl = (e) => {
      const userUrl = e.target.value;
      if (validateUrl(userUrl)) {
        setInvalidUrl(undefined);
        setUrl({ Description: userUrl, Url: userUrl});
      } else {
        setUrl({ Description: userUrl, Url: userUrl});
        setInvalidUrl('Provide a valid URL');
      }
    };

    const setAppFwCategory = (e) => {
      setAppFw(!appFw);
    };

    const closeModal = () => {
      setShowWarning(false);
    };

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
              { Id && showLoadingPeople && <div className={ styles.loadingIndicator }></div> }
          </p>
          <div className= { styles.flex }>
            <p className={ styles.title }>Dynatrace App?</p>
            <p className={ styles.paddingTop }><Toggle label="" onText="Yes" offText="No" onChange={setAppFwCategory} defaultChecked={appFw}/></p>
          </div>
          <div className= { styles.flex } style={{marginTop: '-5px'}}><i>Want to know more about <a href='https://dynatrace.sharepoint.com/sites/Inno_Days/SitePages/Platform-Apps.aspx'>&nbsp;Dynatrace Apps?</a></i></div>
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
            <>
              <DefaultButton className={ styles.cancel } text="Cancel" allowDisabledFocus onClick={cancelUpdate}/>
              <PrimaryButton
                text="Update team"
                className={ styles.save }
                allowDisabledFocus
                style={{backgroundColor:'#0078d4', border: 'none'}}
                onClick={updateTeam}
                disabled={disabled}
              />
            </>
          }
        </div>
        <Dialog 
          isOpen={showWarning} 
          type={DialogType.close} 
          onDismiss={closeModal}
          title='You are about to go on free style mode' 
          subText='' 
          isBlocking={false} 
          closeButtonAriaLabel='Close'  
          maxWidth={'500px'}
        > 
          <p>
            In order to include your project in voting mode your project has to include top 6 members.
            Otherwise, your project will still participate in the event but in free style.
            <br></br>
          </p>
          <DialogFooter> 
            <PrimaryButton
                text="Got it!"
                allowDisabledFocus
                onClick={closeModal}
                disabled={disabled}
              />
          </DialogFooter> 
        </Dialog>
      </>
    );
  };

export default TeamForm;
