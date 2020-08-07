import { Component, OnInit } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { User } from './User';
import { Group } from './Group';
const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';
@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.css']
})

export class DashboardComponent implements OnInit {

  constructor(private http: HttpClient) { }

  lstUsers: User[] = [];
  usrsCnt: number;
  lstGroups: Group[] = [];
  grpName: string;
  grpDesc: string;
  selectedUsers: [] = [];
  profile: any;

  ngOnInit() {
    this.listUsers();
    this.listGroups();
    this.getProfile();
  }

  getProfile() {
    this.http.get(GRAPH_ENDPOINT + '/me')
      .toPromise().then(profile => {
        this.profile = profile;
        console.log('profile', profile);
      });
  }

  listUsers() {
    // ltts group id b383a0fe-55b7-44a1-a0d4-c72fcf9095aa
    // https://graph.microsoft.com/v1.0/teams/b383a0fe-55b7-44a1-a0d4-c72fcf9095aa/channels
    this.http.get('https://graph.microsoft.com/v1.0/users')
      .toPromise().then(res => {
        let result: any = res;
        let list = result.value;
        let i: number = 0;
        list.forEach(itm => {
          let user: User = new User();
          user.id = itm.id;
          user.givenName = itm.givenName;
          user.surname = itm.surname;
          user.userPrincipalName = itm.userPrincipalName;
          this.lstUsers.push(user);
          i++;
        })
        this.usrsCnt = i;
        console.log(this.lstUsers);
      });
  }

  listGroups() {
    this.http.get('https://graph.microsoft.com/v1.0/groups')
      .toPromise().then(res => {
        let result: any = res;
        let list = result.value;
        console.log('list', list);
        list.forEach(itm => {
          let group: Group = new Group();
          group.id = itm.id;
          group.description = itm.description;
          group.displayName = itm.displayName;
          this.lstGroups.push(group);
        })
        console.log(result);
      });
  }

  createGroup() {
    console.log(this.selectedUsers)

    let members_odata_bind: string[] = [];
    let membersExist: boolean = false;
    // collect the member ids to be added
    for (let i = 0; i < this.usrsCnt; i++) {
      if (this.selectedUsers[i] && this.selectedUsers[i] == true) {
        console.log(this.lstUsers[i].id + ' = ' + this.lstUsers[i].givenName);
        members_odata_bind.push('https://graph.microsoft.com/v1.0/users/' + this.lstUsers[i].id)
        membersExist = true;
      }
    }
    if (membersExist == false) {
      alert('select few memberst to create group!')
      return;
    }
    if (this.grpName.trim().length == 0) {
      alert('Group name can not be empty');
      return;
    }
    let owners_odata_bind: string[] = [];
    owners_odata_bind.push('https://graph.microsoft.com/v1.0/users/' + this.profile.id);

    //  sample format form https://docs.microsoft.com/en-us/graph/teams-create-group-and-team
    //   {
    //     "displayName":"Flight 157",
    //     "mailNickname":"flight157",
    //     "description":"Everything about flight 157",
    //     "visibility":"Private",
    //     "groupTypes":["Unified"],
    //     "mailEnabled":true,
    //     "securityEnabled":false,
    //     "members@odata.bind":[
    //         "https://graph.microsoft.com/v1.0/users/bec05f3d-a818-4b58-8c2e-2b4e74b0246d",
    //         "https://graph.microsoft.com/v1.0/users/ae67a4f4-2308-4522-9021-9f402ff0fba8",
    //         "https://graph.microsoft.com/v1.0/users/eab978dd-35d0-4885-8c46-891b7d618783",
    //         "https://graph.microsoft.com/v1.0/users/6a1272b5-f6fc-45c4-95fe-fe7c5a676133"
    //     ],
    //     "owners@odata.bind":[
    //         "https://graph.microsoft.com/v1.0/users/6a1272b5-f6fc-45c4-95fe-fe7c5a676133",
    //         "https://graph.microsoft.com/v1.0/users/eab978dd-35d0-4885-8c46-891b7d618783"
    //     ]
    // }
    // 
    let groupRec = {
      "displayName": this.grpName,
      "mailNickname": this.grpName.trim().replace(' ', ''),
      "description": this.grpDesc,
      "visibility": "Public",
      "groupTypes": ["Unified"],
      "mailEnabled": true,
      "securityEnabled": false,
      "members@odata.bind":
        members_odata_bind
      ,
      "owners@odata.bind":
        owners_odata_bind

    }
    console.log(groupRec);
    this.http.post('https://graph.microsoft.com/v1.0/groups', groupRec).subscribe(res => {
      console.log(res);
    })




  }
}
