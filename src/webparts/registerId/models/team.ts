export interface ITeam {
	Title?: string;
	Id: number;
	Description: string;
	MembersId: number[];
	Project_x0020_link: { Description: string; Url: string };
}

export interface ISPListTeams {
	value: ITeam[];
}
