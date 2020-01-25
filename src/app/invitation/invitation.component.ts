import { Component, OnInit } from "@angular/core";
import { Client, ClientOptions, AuthProvider } from "@microsoft/microsoft-graph-client";
import { MyAuthenticationProvider } from "../MyAuthenticationProvider";
import { BroadcastService, MsalService } from "@azure/msal-angular";
import * as Msal from "msal";
import { HttpClient } from "@angular/common/http";
import { HttpServiceHelper } from "../common/HttpServiceHelper";
import { Subscription } from "rxjs/Subscription";
import { NgForm } from "@angular/forms";

@Component({
  selector: 'app-invitation',
  templateUrl: './invitation.component.html',
  styleUrls: ['./invitation.component.css']
})
export class InvitationComponent implements OnInit {
  private subscription: Subscription;
  userData;
  url = 'https://graph.microsoft.com/v1.0/me';
  name: string;
  emailAddress: string;
  message: string;
  private token = '';


  constructor(
    private authService: MsalService,
    private http: HttpClient,
    private httpService: HttpServiceHelper,
    private broadcastService: BroadcastService
  ) {}

  ngOnInit() {
    // this.subscription =  this.broadcastService.subscribe('msal:acquireTokenFailure', (payload) => {
    //   console.log("acquireTokenFailure" , payload);
    // });

    // this.onSubmit();

    this.subscription = this.broadcastService.subscribe(
      'msal:acquireTokenSuccess',
      payload => {
        console.log('acquire token success' + JSON.stringify(payload));
      }
    );

    // will work for acquireTokenSilent and acquireTokenPopup
    this.subscription = this.broadcastService.subscribe(
      'msal:acquireTokenFailure',
      payload => {
        console.log('acquire token failure ' + JSON.stringify(payload));
        // if (
        //   payload.indexOf("consent_required") !== -1 ||
        //   payload.indexOf("interaction_required") != -1
        // ) {
        // }

        this.authService
          .acquireTokenPopup([
            'User.Invite.All',
            'User.ReadWrite.All',
            'Directory.ReadWrite.All'
            // ,"Mail.Send"
          ])
          .then(
            token => {
              console.log('token', token);
              this.token = token;
              // this.getUSerProfile();
              // this.onSubmit();
            },
            error => {}
          );
      }
    );

    this.getUSerProfile();
  }

  getUSerProfile() {
    this.httpService.httpGetRequest(this.url).subscribe(
      data => {
        this.userData = data;
      },
      error => {
        console.error(
          ' Http get request to MS Graph failed' + JSON.stringify(error)
        );
      }
    );

    const accessTokenRequest = {
      scopes: [
        'user.invite.all',
        'user.readWrite.all',
        'directory.readWrite.all'

      ]
  }

    this.authService
      .acquireTokenSilent(accessTokenRequest.scopes)
      .then(
        token => {
          console.log('token', token);
          this.token = token;
          // this.getUSerProfile();
          // this.onSubmit();
        },
        error => {}
      );


  }

  // tslint:disable-next-line: use-lifecycle-interface
  ngOnDestroy() {
    this.broadcastService.getMSALSubject().next(1);
    if (this.subscription) {
      this.subscription.unsubscribe();
    }
  }

  async onSubmit() {
    console.log('submitted');

    const clientOptions: ClientOptions = {
      authProvider: new MyAuthenticationProvider(this.token)
    };
    const client = Client.initWithMiddleware(clientOptions);

    const invitation = {
      invitedUserDisplayName: this.name,
      invitedUserEmailAddress: this.emailAddress,
      inviteRedirectUrl: 'http://localhost:4200',
      sendInvitationMessage: true,
      customizedMessageBody: this.message
    };
    const res = await client
      .api('/invitations')
      .post(invitation);
    console.log(res);
  }
}

    // const options = {
    //   authProvider,
    // };

    // const client = Client.init(options);

    // const invitation = {
    //   invitedUserEmailAddress: "yyy@test.com",
    //   inviteRedirectUrl: "https://myapp.com"
    // };

    // let res = await client.api('/invitations')
    //   .post({invitation : invitation});
