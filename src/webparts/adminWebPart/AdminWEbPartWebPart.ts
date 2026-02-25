import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import AdminWEbPart from './components/AdminWebPart';
// import { IAdminWEbPartProps } from './components/IAdminWEbPartProps';
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

export interface IAdminWEbPartWebPartProps {
  description: string;
}

export default class AdminWEbPartWebPart extends BaseClientSideWebPart<IAdminWEbPartWebPartProps> {
  private _sp: SPFI;

   public async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

   public render(): void {
    const element= React.createElement(
      AdminWEbPart,
      {
        sp: this._sp
      }
    );

    ReactDom.render(element, this.domElement);
  }
}
