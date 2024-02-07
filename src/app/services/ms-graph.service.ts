import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import { MsalService } from '@azure/msal-angular';

interface TeamCreationRequest {
  "template@odata.bind": string;
  displayName: string;
  description: string;
  [propertyName: string]: any; // Permite propiedades adicionales de cualquier tipo
}

@Injectable({
  providedIn: 'root'
})

export class MsGraphService {

  private graphClient: Client;

  constructor(private authService: MsalService) {
    this.graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          return this.getToken()
        }
      }
    });
  }
  
  async createTeam(teamName: string): Promise<any> {
    const group = {
      displayName: teamName,
      description: `Grupo de Office 365 para ${teamName}`,
      mailNickname: teamName.replace(/\s+/g, '').toLowerCase(),
      securityEnabled: false,
      mailEnabled: true,
      groupTypes: ['Unified']
    };
  
    try {
      // Crear el grupo de Office 365
      const createdGroup = await this.graphClient.api('/groups').post(group);
  
      // Preparar el objeto de solicitud para crear el equipo
      let teamCreationRequest: TeamCreationRequest = {
        "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        displayName: teamName,
        description: `Team para ${teamName}`,
      };
  
      // Convertir el grupo en un equipo
      const createdTeam = await this.graphClient.api('/teams').post(teamCreationRequest);
  
      return createdTeam;
    } catch (error) {
      console.error("Error creando el team:", error);
      throw error;
    }
  }
    

  async createChannel(teamId: string, channelName: string): Promise<any> {
    const channel = {
      displayName: channelName,
      description: `Canal de ${channelName} en el equipo`,
      membershipType: "standard", // O "private" para canales privados
    };
  
    try {
      const createdChannel = await this.graphClient.api(`/teams/${teamId}/channels`).post(channel);
      return createdChannel;
    } catch (error) {
      console.error("Error creando el canal:", error);
      throw error;
    }
  }
  

  async getToken(): Promise<string> {
    const request: any = {
      scopes: ['user.read', 'Chat.ReadWrite', 'Team.Create'],
      account: this.authService.instance.getActiveAccount()
    };
  
    const accessToken = await this.authService.acquireTokenSilent(request).toPromise()
    console.log("Token =>", accessToken?.accessToken);
    
    return accessToken!.accessToken
  }
}
