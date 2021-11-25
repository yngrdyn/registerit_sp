export interface ITeam {
	Title?: string;
	Id: number;
	Description: string;
	MembersId: number[];
	Project_x0020_link: { Description: string; Url: string };
	AppFw: boolean;
	Recruiting: boolean;
}

export interface ISPListTeams {
	value: ITeam[];
}
