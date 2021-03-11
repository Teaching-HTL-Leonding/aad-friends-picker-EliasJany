import { HttpClient } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { AuthenticationResult } from '@azure/msal-browser';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"; 

@Component({
  selector: 'app-root',
  templateUrl: 'app.component.html',
  styles: []
})
export class AppComponent implements OnInit{
  loggedIn = false;  
  profile?: MicrosoftGraph.User;

  constructor(private authService: MsalService, private client: HttpClient) {}

  ngOnInit(): void {
    this.checkAccount();
  }

  checkAccount() {
    this.loggedIn = this.authService.instance.getAllAccounts().length > 0;
  }

  login() {
    this.authService
      .loginPopup()
      .subscribe((response: AuthenticationResult) => {
        this.authService.instance.setActiveAccount(response.account);
        this.checkAccount();
      });
  }

  logout() {
    this.authService.logout();
  }

  getProfile(){
    this.client.get('https://graph.microsoft.com/v1.0/me')
    .subscribe(profile => this.profile = profile);
  }
}
