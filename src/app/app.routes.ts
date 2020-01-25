import { HomeComponent } from './home/home.component';
import { ErrorComponent } from './error.component';
// import { AppComponent } from './app.component';
import { InvitationComponent } from './invitation/invitation.component';
import { Routes } from '@angular/router';
import {MsalGuard} from "@azure/msal-angular";


export const appRoutes: Routes = [
  { path: 'home', component: HomeComponent  },
  { path: 'invitation', component: InvitationComponent , canActivate : [MsalGuard] },
  { path: '', redirectTo: 'app', pathMatch: 'full' },
  { path: '**', component: ErrorComponent }
];




