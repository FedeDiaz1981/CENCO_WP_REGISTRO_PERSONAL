import { SPFI } from "@pnp/sp";

export interface IWpRegistroPersonalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: SPFI;
  siteUrl: string;
  listaPersonal: string;
  vistaModificar: string;
  vistaVisualizar: string;
  vistaDarBaja: string;
  mostrarIngresar: boolean;
  mostrarModificar: boolean;
  mostrarVisualizar: boolean;
  mostrarDarBaja: boolean;
  filtrarPorProveedor: boolean;
  borrar: boolean;
  bloquearEmpresa: boolean; // ✅ NUEVO
}
