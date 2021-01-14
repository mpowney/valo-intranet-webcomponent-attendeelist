import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent } from '@valo/extensibility';
import { AttendeeRegister } from '.';

const provisionListModalAlreadyLoadedObjectName: string = "attendeeRegisterProvisionListModalLoaded";

export class AttendeeRegisterWebComponent extends BaseWebComponent {

  public constructor() {
    super();
  }

  public async connectedCallback() {

    (window as any)[provisionListModalAlreadyLoadedObjectName] = (window as any)[provisionListModalAlreadyLoadedObjectName] || 0;
    let props: any = this.resolveAttributes();
    // You can use this._ctx here to access current Web Part context
    props._ctx = this._ctx;
    const attendeeRegister = <AttendeeRegister {...props} />;
    ReactDOM.render(attendeeRegister, this);
    // console.log(`connectedCallback() called`);

  }

  public disconnectedCallback() {
    (window as any)[provisionListModalAlreadyLoadedObjectName]--;
    if ((window as any)[provisionListModalAlreadyLoadedObjectName] <= 0) delete (window as any)[provisionListModalAlreadyLoadedObjectName];
    // console.log(`disconnectedCallback() called`);

  }

}
