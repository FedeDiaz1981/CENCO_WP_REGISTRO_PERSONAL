import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import RegistroPersonal, { IRegistroPersonalProps } from './components/WpRegistroPersonal';

import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IWP_CENCO_Registro_PersonalWebPartProps {
  filtrarPorProveedor: boolean;  // <-- Toggle para filtrar grilla por proveedor
}

export default class WP_CENCO_Registro_PersonalWebPart
  extends BaseClientSideWebPart<IWP_CENCO_Registro_PersonalWebPartProps> {

  private _sp!: SPFI;

  public async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    const element: React.ReactElement<IRegistroPersonalProps> = React.createElement(
      RegistroPersonal,
      {
        sp: this._sp,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        filtrarPorProveedor: this.properties.filtrarPorProveedor // <-- Se pasa al componente
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configuración' },
          groups: [
            {
              groupName: 'Opciones de la grilla',
              groupFields: [
                PropertyPaneToggle('filtrarPorProveedor', {
                  label: 'Filtrar registros por proveedor del usuario',
                  onText: 'Sí',
                  offText: 'No'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
