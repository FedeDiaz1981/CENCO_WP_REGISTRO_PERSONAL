import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import RegistroPersonal, { IRegistroPersonalProps } from './components/WpRegistroPersonal';

import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/views/list';
import '@pnp/sp/items';

export interface IWP_CENCO_Registro_PersonalWebPartProps {
  listaPersonal: string;
  correosDefault: string;
  listaAdjuntos: string;
  redirigir: boolean;
  urlRedireccion: string;
  vistaModificar: string;
  vistaVisualizar: string;
  vistaDarBaja: string;
  mostrarIngresar: boolean;
  mostrarModificar: boolean;
  mostrarVisualizar: boolean;
  mostrarDarBaja: boolean;
  filtrarPorProveedor: boolean;
  borrar: boolean;
  bloquearEmpresa: boolean;
}

type ViewInfo = {
  Id: string;
  Title: string;
  DefaultView?: boolean;
  Hidden?: boolean;
};

export default class WP_CENCO_Registro_PersonalWebPart
  extends BaseClientSideWebPart<IWP_CENCO_Registro_PersonalWebPartProps> {
  private _sp!: SPFI;
  private _viewOptions: IPropertyPaneDropdownOption[] = [];
  private readonly _defaultListaPersonal = 'Personal';

  public async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    this.properties.listaPersonal = this.properties.listaPersonal?.trim() || this._defaultListaPersonal;
    this.properties.mostrarIngresar = this.properties.mostrarIngresar ?? true;
    this.properties.mostrarModificar = this.properties.mostrarModificar ?? true;
    this.properties.mostrarVisualizar = this.properties.mostrarVisualizar ?? true;
    this.properties.mostrarDarBaja = this.properties.mostrarDarBaja ?? true;
    this.properties.redirigir = this.properties.redirigir ?? false;
    await this._cargarVistasConfiguradas();
  }

  private _syncSelectedViews(): void {
    const availableIds = new Set(this._viewOptions.map((opt) => String(opt.key)));
    const fallback = this._viewOptions[0] ? String(this._viewOptions[0].key) : '';

    if (!this.properties.vistaModificar || !availableIds.has(this.properties.vistaModificar)) {
      this.properties.vistaModificar = fallback;
    }

    if (!this.properties.vistaVisualizar || !availableIds.has(this.properties.vistaVisualizar)) {
      this.properties.vistaVisualizar = fallback;
    }

    if (!this.properties.vistaDarBaja || !availableIds.has(this.properties.vistaDarBaja)) {
      this.properties.vistaDarBaja = fallback;
    }
  }

  private async _cargarVistasConfiguradas(): Promise<void> {
    const listTitle = this.properties.listaPersonal?.trim() || this._defaultListaPersonal;

    try {
      const vistas = (await this._sp.web.lists
        .getByTitle(listTitle)
        .views.select('Id', 'Title', 'DefaultView', 'Hidden')()) as ViewInfo[];

      this._viewOptions = vistas
        .filter((vista) => !vista.Hidden)
        .sort((a, b) => {
          if (a.DefaultView && !b.DefaultView) return -1;
          if (!a.DefaultView && b.DefaultView) return 1;
          return a.Title.localeCompare(b.Title, 'es');
        })
        .map((vista) => ({
          key: vista.Id,
          text: vista.Title
        }));

      this._syncSelectedViews();
    } catch {
      this._viewOptions = [];
      this.properties.vistaModificar = '';
      this.properties.vistaVisualizar = '';
      this.properties.vistaDarBaja = '';
    }

    this.context.propertyPane.refresh();
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listaPersonal' && oldValue !== newValue) {
      this.properties.listaPersonal = String(newValue || '').trim();
      void this._cargarVistasConfiguradas();
    }

    if (propertyPath === 'redirigir') {
      this.context.propertyPane.refresh();
    }
  }

  public render(): void {
    const element: React.ReactElement<IRegistroPersonalProps> = React.createElement(
      RegistroPersonal,
      {
        sp: this._sp,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        listaPersonal: this.properties.listaPersonal || this._defaultListaPersonal,
        correosDefault: this.properties.correosDefault || '',
        listaAdjuntos: this.properties.listaAdjuntos || '',
        redirigir: this.properties.redirigir ?? false,
        urlRedireccion: this.properties.urlRedireccion || '',
        vistaModificar: this.properties.vistaModificar || '',
        vistaVisualizar: this.properties.vistaVisualizar || '',
        vistaDarBaja: this.properties.vistaDarBaja || '',
        mostrarIngresar: this.properties.mostrarIngresar ?? true,
        mostrarModificar: this.properties.mostrarModificar ?? true,
        mostrarVisualizar: this.properties.mostrarVisualizar ?? true,
        mostrarDarBaja: this.properties.mostrarDarBaja ?? true,
        filtrarPorProveedor: this.properties.filtrarPorProveedor,
        borrar: this.properties.borrar,
        bloquearEmpresa: this.properties.bloquearEmpresa
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
          header: { description: 'Configuracion' },
          groups: [
            {
              groupName: 'Origen de datos',
              groupFields: [
                PropertyPaneTextField('listaPersonal', {
                  label: 'Lista destino',
                  value: this.properties.listaPersonal || this._defaultListaPersonal,
                  placeholder: 'Escribi el nombre exacto de la lista'
                }),
                PropertyPaneTextField('correosDefault', {
                  label: 'Correos por default',
                  value: this.properties.correosDefault || '',
                  placeholder: 'correo1@dominio.com; correo2@dominio.com'
                }),
                PropertyPaneTextField('listaAdjuntos', {
                  label: 'Lista de adjuntos',
                  value: this.properties.listaAdjuntos || '',
                  placeholder: 'Documentacion'
                }),
                PropertyPaneToggle('redirigir', {
                  label: 'Redirigir despues de crear',
                  onText: 'Si',
                  offText: 'No'
                })
              ]
            },
            {
              groupName: 'Vistas de la grilla',
              groupFields: [
                PropertyPaneDropdown('vistaModificar', {
                  label: 'Vista para Modificar',
                  options: this._viewOptions,
                  selectedKey: this.properties.vistaModificar || undefined,
                  disabled: this._viewOptions.length === 0
                }),
                PropertyPaneDropdown('vistaVisualizar', {
                  label: 'Vista para Visualizar',
                  options: this._viewOptions,
                  selectedKey: this.properties.vistaVisualizar || undefined,
                  disabled: this._viewOptions.length === 0
                }),
                PropertyPaneDropdown('vistaDarBaja', {
                  label: 'Vista para Dar de baja',
                  options: this._viewOptions,
                  selectedKey: this.properties.vistaDarBaja || undefined,
                  disabled: this._viewOptions.length === 0
                })
              ]
            },
            {
              groupName: 'Visibilidad del formulario',
              groupFields: [
                PropertyPaneToggle('mostrarIngresar', {
                  label: 'Mostrar opcion Ingresar',
                  onText: 'Si',
                  offText: 'No'
                }),
                PropertyPaneToggle('mostrarModificar', {
                  label: 'Mostrar opcion Modificar',
                  onText: 'Si',
                  offText: 'No'
                }),
                PropertyPaneToggle('mostrarVisualizar', {
                  label: 'Mostrar opcion Visualizar',
                  onText: 'Si',
                  offText: 'No'
                }),
                PropertyPaneToggle('mostrarDarBaja', {
                  label: 'Mostrar opcion Dar de baja',
                  onText: 'Si',
                  offText: 'No'
                })
              ]
            },
            {
              groupName: 'Opciones de comportamiento',
              groupFields: [
                PropertyPaneToggle('filtrarPorProveedor', {
                  label: 'Filtrar registros por proveedor del usuario',
                  onText: 'Si',
                  offText: 'No'
                }),
                PropertyPaneToggle('borrar', {
                  label: 'Borrar registro al dar de baja',
                  onText: 'Si (eliminar registro)',
                  offText: 'No (marcar inactivo)'
                }),
                PropertyPaneToggle('bloquearEmpresa', {
                  label: 'Bloquear empresa segun proveedor del usuario',
                  onText: 'Si (bloqueado)',
                  offText: 'No (editable)'
                }),
                ...(this.properties.redirigir
                  ? [
                      PropertyPaneTextField('urlRedireccion', {
                        label: 'URL de redireccion',
                        value: this.properties.urlRedireccion || '',
                        placeholder: 'https://...'
                      })
                    ]
                  : [])
              ]
            }
          ]
        }
      ]
    };
  }
}
