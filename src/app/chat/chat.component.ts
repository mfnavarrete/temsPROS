import { Component, OnInit } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import { MsGraphService } from '../services/ms-graph.service';

@Component({
  selector: 'app-chat',
  templateUrl: './chat.component.html',
  styleUrls: ['./chat.component.css']
})
export class ChatComponent implements OnInit {

  teamsID: string = ""
  channelsID: string = ""

  constructor(private msGraphService: MsGraphService) { }

  ngOnInit(): void {
  }

  createTeam() {
    const teamName: string = "New Team PROS"
    this.msGraphService.createTeam(teamName).then(team => {
      console.log("Created Team: ", team);
    })
  }

  getToken() {
    this.msGraphService.getToken()
  }

  getAuthenticatedClient(accessToken: string) {
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken); // Primero el error, luego el token
      }
    });
  
    return client;
  }
  
  async sendMessage(accessToken: string) {
    const client = this.getAuthenticatedClient(accessToken);
    
    const message = {
      body: {
        content: "Hello World!"
      }
    };
  
    try {
      await client.api(`/teams/${this.teamsID}/channels/${this.channelsID}/messages`)
        .post(message);
    } catch (error) {
      console.error(error);
    }
  }

}
