import { IWebComponentProvider, IWebComponentDefinition } from '@valo/extensibility';
import { AttendeeRegisterWebComponent } from './components';

export class CustomWebComponentsLibraryLibrary implements IWebComponentProvider {
  public name(): string {
    return 'CustomWebComponentsLibraryLibrary';
  }

  /**
   * Return your custom web components
   */
  public getWebComponents(): IWebComponentDefinition<any>[] {
    return [
      {
        name: 'attendee-register',
        class: AttendeeRegisterWebComponent
      }
    ];
  }


}
