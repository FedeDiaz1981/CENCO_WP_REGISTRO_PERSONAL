import { SPFI } from "@pnp/sp";

export interface IWpRegistroPersonalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: SPFI;
  siteUrl: string;
  filtrarPorProveedor: boolean;
  borrar: boolean;
  bloquearEmpresa: boolean; // ✅ NUEVO
}