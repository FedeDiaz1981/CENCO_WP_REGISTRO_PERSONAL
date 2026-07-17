import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
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
  filtrarPorProveedor: boolean;
  borrar: boolean;
  bloquearEmpresa: boolean;
  listaPersonal: string;
}

export default class WP_CENCO_Registro_PersonalWebPart
  extends BaseClientSideWebPart<IWP_CENCO_Registro_PersonalWebPartProps> {

  private _sp!: SPFI;
  private _listaPersonalOptions: IPropertyPaneDropdownOption[] = [];
  private readonly _defaultListaPersonal = 'Personal';

  public async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    await this._cargarListasDestino();
  }

  private async _cargarListasDestino(): Promise<void> {
    try {
      const listas = await this._sp.web.lists
        .select('Title', 'Hidden', 'BaseTemplate')();

      this._listaPersonalOptions = listas
        .filter((lista: any) => !lista.Hidden && lista.BaseTemplate === 100)
        .map((lista: any) => ({
          key: lista.Title,
          text: lista.Title
        }))
        .sort((a, b) => a.text.localeCompare(b.text, 'es'));

      if (!this._listaPersonalOptions.length) {
        this._listaPersonalOptions = [
          {
            key: this._defaultListaPersonal,
            text: this._defaultListaPersonal
          }
        ];
      }

      const listaActual = this.properties.listaPersonal?.trim();
      const existeActual = listaActual
        ? this._listaPersonalOptions.some((opt) => opt.key === listaActual)
        : false;

      if (!listaActual || !existeActual) {
        this.properties.listaPersonal =
          String(this._listaPersonalOptions[0].key) || this._defaultListaPersonal;
      }
    } catch {
      this._listaPersonalOptions = [
        {
          key: this._defaultListaPersonal,
          text: this._defaultListaPersonal
        }
      ];

      if (!this.properties.listaPersonal) {
        this.properties.listaPersonal = this._defaultListaPersonal;
      }
    }

    this.context.propertyPane.refresh();
  }

  public render(): void {
    const element: React.ReactElement<IRegistroPersonalProps> = React.createElement(
      RegistroPersonal,
      {
        sp: this._sp,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        filtrarPorProveedor: this.properties.filtrarPorProveedor,
        borrar: this.properties.borrar,
        bloquearEmpresa: this.properties.bloquearEmpresa,
        listaPersonal: this.properties.listaPersonal || this._defaultListaPersonal
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
                PropertyPaneDropdown('listaPersonal', {
                  label: 'Lista destino',
                  options: this._listaPersonalOptions,
                  selectedKey: this.properties.listaPersonal || this._defaultListaPersonal,
                  disabled: this._listaPersonalOptions.length === 0
                }),
                PropertyPaneToggle('filtrarPorProveedor', {
                  label: 'Filtrar registros por proveedor del usuario',
                  onText: 'Sí',
                  offText: 'No'
                }),
                PropertyPaneToggle('borrar', {
                  label: 'Borrar registro al dar de baja',
                  onText: 'Sí (eliminar registro)',
                  offText: 'No (marcar inactivo)'
                }),
                PropertyPaneToggle('bloquearEmpresa', {
                  label: 'Bloquear empresa (según proveedor del usuario)',
                  onText: 'Sí (bloqueado)',
                  offText: 'No (editable)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
