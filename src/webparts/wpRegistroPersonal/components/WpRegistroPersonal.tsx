import * as React from "react";
import { SPFI } from "@pnp/sp";
import {
  ThemeProvider,
  createTheme,
  Stack,
  StackItem,
  Label,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  DatePicker,
  DayOfWeek,
  MessageBar,
  MessageBarType,
  ProgressIndicator,
  Icon,
  DetailsList,
  IColumn,
  SelectionMode,
  Spinner,
  SpinnerSize,
  Selection,
  SearchBox,
  Dialog,
  DialogType,
  DialogFooter,
  IButtonStyles,
  IMessageBarStyles,
} from "@fluentui/react";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/fields/list";

export interface IRegistroPersonalProps {
  sp: SPFI;
  siteUrl: string;
  filtrarPorProveedor: boolean;
  borrar: boolean; // toggle de la webpart
  bloquearEmpresa: boolean; // ✅ NUEVO: si true bloquea y autodetecta, si false deja elegir
}

type Modo = "Ingresar" | "Modificar" | "Dar de baja";

interface PersonaForm {
  Documento: string;
  Nombre: string;
  ApellidoPaterno: string;
  ApellidoMaterno: string;
  TipoDocumento?: string;
  Puesto?: string;
  Especificar?: string;
  Licencia?: string;
  Categoria?: string;
  correosnotificacion?: string;
  CorreosNotificacion?: string;
}

// -------- Opciones --------
const opcionesTipoDocumento: IDropdownOption[] = [
  { key: "DNI", text: "DNI" },
  { key: "Pasaporte", text: "Pasaporte" },
  { key: "Carnet", text: "Carnet" },
];

const opcionesPuesto: IDropdownOption[] = [
  { key: "Conductor", text: "Conductor" },
  { key: "Auxiliar de carga", text: "Auxiliar de carga" },
  { key: "Coordinador de operación", text: "Coordinador de operación" },
  { key: "Otro", text: "Otro" },
];

const opcionesCategoria: IDropdownOption[] = [
  { key: "A", text: "A" },
  { key: "B", text: "B" },
  { key: "C", text: "C" },
];

const getDocumentoLengthRequerido = (tipo?: string): number | undefined => {
  if (tipo === "DNI") return 8;
  if (tipo === "Pasaporte" || tipo === "Carnet") return 9;
  return undefined;
};

const LST_PERSONAS = "Personal";
const LST_DOCS = "Documentacion";
const LST_PROVEEDORES = "Proveedores";

// tokens de Stack
const stackTokens: { childrenGap: number } = { childrenGap: 12 };
const GRID_PAGE_SIZE = 10;

const esc = (s: string) => s.replace(/'/g, "''");
const dateToISO = (d?: Date | null) => (d ? d.toISOString() : null);

// helper para limpiar HTML del campo de correos
const stripHtml = (html?: string | null): string =>
  html ? html.replace(/<[^>]*>/g, "").trim() : "";

type DocFields = { Caducidad?: string | null; Emision?: string | null };

// ===== Tema Cencosud =====
const theme = createTheme({
  palette: {
    themePrimary: "#005596",
    themeLighterAlt: "#f2f7fb",
    themeLighter: "#deebf8",
    themeLight: "#c2daf1",
    themeTertiary: "#7eb2db",
    themeSecondary: "#2f7fc0",
    themeDarkAlt: "#004d87",
    themeDark: "#00406f",
    themeDarker: "#002f51",
    neutralLighterAlt: "#f4f9ff",
    neutralLighter: "#edf4fb",
    neutralLight: "#d7e5f3",
    neutralQuaternaryAlt: "#ccdaea",
    neutralQuaternary: "#c1d3e6",
    neutralTertiaryAlt: "#b5c7dc",
    neutralTertiary: "#333333",
    neutralSecondary: "#55687c",
    neutralPrimaryAlt: "#233140",
    neutralPrimary: "#1e2a36",
    neutralDark: "#1f1f1f",
    black: "#1a1a1a",
    white: "#ffffff",
  },
  effects: {
    roundedCorner2: "18px",
    elevation8: "0 12px 28px rgba(0,87,166,.12)" as any,
  },
});

const BRAND = {
  canvas: "#eef4fb",
  shell: "#f7fbff",
  ink: "#1e2a36",
  muted: "#617284",
  border: "#cad9ea",
};

const HERO_BG =
  "radial-gradient(circle at 16% 18%, rgba(255,255,255,.18) 0 72px, transparent 73px), radial-gradient(circle at 86% -12%, rgba(255,255,255,.18) 0 118px, transparent 120px), linear-gradient(135deg, #005596 0%, #0067b2 48%, #0072bc 100%)";

const pageShellStyles = {
  root: {
    maxWidth: 1040,
    margin: "0 auto",
    padding: 18,
    background: `linear-gradient(180deg, ${BRAND.shell} 0%, ${BRAND.canvas} 100%)`,
    borderRadius: 32,
    border: `1px solid ${BRAND.border}`,
    boxShadow: "0 24px 54px rgba(0,87,166,.12)",
  },
};

const heroPanelStyles = {
  root: {
    background: HERO_BG,
    borderRadius: 26,
    padding: 24,
    boxShadow: "0 22px 42px rgba(0,87,166,.22)",
    width: "100%",
    overflow: "hidden" as const,
  },
};

const sectionCardStyles = {
  root: {
    background: "rgba(255,255,255,.92)",
    borderRadius: 24,
    padding: 20,
    border: `1px solid ${BRAND.border}`,
    boxShadow: "0 16px 32px rgba(0,87,166,.08)",
  },
};

const sectionTitleStyles = {
  root: {
    display: "inline-flex" as const,
    alignSelf: "flex-start" as const,
    padding: "8px 16px",
    borderRadius: 999,
    background: "linear-gradient(135deg, #005596 0%, #0072bc 100%)",
    color: "#ffffff",
    fontSize: 14,
    fontWeight: 700 as const,
    lineHeight: 1.2,
    boxShadow: "0 10px 18px rgba(0,87,166,.18)",
    marginBottom: 6,
  },
};

const MESSAGE_BAR_FONT_FAMILY = "\"Segoe UI\", Tahoma, Geneva, Verdana, sans-serif";

const messageBarStyles: IMessageBarStyles = {
  root: {
    borderRadius: 2,
    overflow: "hidden",
    boxShadow: "none",
  },
  content: {
    padding: "12px 16px",
  },
  iconContainer: {
    alignSelf: "flex-start",
    paddingTop: 2,
  },
  icon: {
    fontSize: 16,
  },
  text: {
    fontSize: 14,
    lineHeight: 1.5,
    fontFamily: MESSAGE_BAR_FONT_FAMILY,
  },
  innerText: {
    fontSize: 14,
    lineHeight: 1.5,
    fontFamily: MESSAGE_BAR_FONT_FAMILY,
    fontWeight: 400,
  },
};

const errorMessageBarStyles: IMessageBarStyles = {
  root: {
    borderRadius: 2,
    overflow: "hidden",
    boxShadow: "none",
    background: "#fde7e9",
    border: "1px solid #f1c5ca",
  },
  content: {
    padding: "12px 16px",
  },
  iconContainer: {
    alignSelf: "flex-start",
    paddingTop: 2,
  },
  icon: {
    fontSize: 16,
    color: "#d13438",
  },
  text: {
    fontSize: 14,
    lineHeight: 1.5,
    fontFamily: MESSAGE_BAR_FONT_FAMILY,
    color: "#323130",
  },
  innerText: {
    fontSize: 14,
    lineHeight: 1.5,
    fontFamily: MESSAGE_BAR_FONT_FAMILY,
    fontWeight: 400,
    color: "#323130",
  },
};

const errorMessageBarIconProps = { iconName: "StatusErrorFull" };

const dangerBannerStyles = {
  root: {
    display: "flex",
    alignItems: "flex-start",
    gap: 12,
    width: "100%",
    padding: "12px 16px",
    boxSizing: "border-box" as const,
    borderRadius: 2,
    border: "1px solid #f1c5ca",
    background: "#fde7e9",
  },
  icon: {
    fontSize: 16,
    color: "#d13438",
    marginTop: 2,
    flexShrink: 0,
  },
  text: {
    flex: 1,
    minWidth: 0,
    color: "#323130",
    fontFamily: MESSAGE_BAR_FONT_FAMILY,
    fontSize: 14,
    lineHeight: "21px",
    fontWeight: 400,
    whiteSpace: "normal" as const,
    wordBreak: "break-word" as const,
  },
};

const infoBannerStyles = {
  root: {
    display: "flex",
    alignItems: "flex-start",
    gap: 12,
    width: "100%",
    padding: "12px 16px",
    boxSizing: "border-box" as const,
    borderRadius: 2,
    border: "1px solid #e1dfdd",
    background: "#f3f2f1",
  },
  icon: {
    fontSize: 16,
    color: "#605e5c",
    marginTop: 2,
    flexShrink: 0,
  },
  text: {
    flex: 1,
    minWidth: 0,
    color: "#323130",
    fontFamily: MESSAGE_BAR_FONT_FAMILY,
    fontSize: 14,
    lineHeight: "21px",
    fontWeight: 400,
    whiteSpace: "normal" as const,
    wordBreak: "break-word" as const,
  },
};

const successBannerStyles = {
  root: {
    display: "flex",
    alignItems: "flex-start",
    gap: 12,
    width: "100%",
    padding: "12px 16px",
    boxSizing: "border-box" as const,
    borderRadius: 2,
    border: "1px solid #b7dfb0",
    background: "#dff6dd",
  },
  icon: {
    fontSize: 16,
    color: "#107c10",
    marginTop: 2,
    flexShrink: 0,
  },
  text: {
    flex: 1,
    minWidth: 0,
    color: "#323130",
    fontFamily: MESSAGE_BAR_FONT_FAMILY,
    fontSize: 14,
    lineHeight: "21px",
    fontWeight: 400,
    whiteSpace: "normal" as const,
    wordBreak: "break-word" as const,
  },
};

const requiredFieldsListStyles = {
  title: {
    fontWeight: 600,
    marginBottom: 6,
  },
  list: {
    margin: 0,
    paddingLeft: 18,
  },
  item: {
    marginBottom: 2,
  },
};

const primaryButtonStyles: IButtonStyles = {
  root: {
    minHeight: 44,
    padding: "0 20px",
    borderRadius: 999,
    border: "none",
    background: "linear-gradient(135deg, #005596 0%, #0072bc 100%)",
    boxShadow: "0 12px 22px rgba(0,87,166,.2)",
  },
  rootHovered: {
    background: "linear-gradient(135deg, #004d87 0%, #0067b2 100%)",
    boxShadow: "0 14px 24px rgba(0,87,166,.24)",
  },
  rootPressed: {
    background: "linear-gradient(135deg, #00406f 0%, #005596 100%)",
  },
  label: {
    fontWeight: 700,
    color: "#ffffff",
  },
  icon: {
    color: "#ffffff",
  },
};

const secondaryButtonStyles: IButtonStyles = {
  root: {
    minHeight: 44,
    padding: "0 20px",
    borderRadius: 999,
    border: `1px solid ${BRAND.border}`,
    background: "rgba(255,255,255,.92)",
    boxShadow: "0 8px 18px rgba(0,87,166,.08)",
  },
  rootHovered: {
    background: "#ffffff",
    borderColor: "#8bb8df",
  },
  label: {
    fontWeight: 600,
    color: BRAND.ink,
  },
};

const modeButtonStyles: IButtonStyles = {
  root: {
    minWidth: 176,
    minHeight: 112,
    padding: "16px 18px",
    borderRadius: 24,
    border: "1px solid rgba(255,255,255,.26)",
    background: "rgba(255,255,255,.14)",
    boxShadow: "0 14px 24px rgba(0,0,0,.12)",
    backdropFilter: "blur(10px)",
    color: "#ffffff",
  },
  rootHovered: {
    background: "rgba(255,255,255,.22)",
    borderColor: "rgba(255,255,255,.34)",
    transform: "translateY(-1px)",
  },
  rootPressed: {
    background: "rgba(255,255,255,.28)",
  },
  rootChecked: {
    background: "#ffffff",
    borderColor: "#ffffff",
    boxShadow: "0 12px 22px rgba(0,0,0,.14)",
    color: "#005596",
  },
  rootCheckedHovered: {
    background: "#ffffff",
    borderColor: "#ffffff",
    boxShadow: "0 12px 22px rgba(0,0,0,.14)",
    color: "#005596",
  },
  flexContainer: {
    flexDirection: "column",
    justifyContent: "center",
    alignItems: "center",
    gap: 10,
  },
  icon: {
    fontSize: 34,
    height: 40,
    lineHeight: "40px",
    margin: 0,
    color: "inherit",
  },
  textContainer: {
    width: "100%",
    display: "flex",
    justifyContent: "center",
  },
  label: {
    width: "100%",
    textAlign: "center",
    fontWeight: 600,
    lineHeight: 1.2,
    whiteSpace: "normal",
    margin: 0,
    color: "inherit",
  },
};

const modeTileWrapStyles = {
  root: {
    position: "relative" as const,
  },
};

const modeTileDotStyles = {
  position: "absolute" as const,
  top: 10,
  right: 10,
  width: 12,
  height: 12,
  borderRadius: "50%",
  border: "2px solid #ffffff",
  background: "#0072bc",
  boxShadow: "0 0 0 1px rgba(0,0,0,.12)",
};

// estilos para que los campos no pisen el borde redondeado del contenedor
const roundedField = {
  fieldGroup: {
    borderRadius: 18,
    minHeight: 44,
    borderColor: BRAND.border,
    background: theme.palette.white,
    boxShadow: "0 4px 14px rgba(0,87,166,.05)",
  },
  field: {
    fontSize: 14,
  },
};
const roundedDropdown = {
  title: {
    borderRadius: 18,
    minHeight: 44,
    lineHeight: 42,
    borderColor: BRAND.border,
    background: theme.palette.white,
    boxShadow: "0 4px 14px rgba(0,87,166,.05)",
  },
};
const roundedDatePicker = {
  textField: {
    selectors: {
      ".ms-TextField-fieldGroup": {
        borderRadius: 18,
        minHeight: 44,
        borderColor: BRAND.border,
        background: theme.palette.white,
        boxShadow: "0 4px 14px rgba(0,87,166,.05)",
      },
    },
  },
};

// ====== Grilla de documentación para MODIFICAR ======
type Attach = { name: string; href: string };

type DocRow = {
  key: string;
  label: string;
  tipo: "cad" | "emi";
  fecha: Date | null;
  attachments: Attach[];
  editing?: boolean;
  justUpdated?: boolean;
  file?: File | null;
};

function toAbs(site: string, rel: string) {
  return site.replace(/\/$/, "") + rel;
}

const DOC_DEFS = [
  { key: "DNI", tipo: "cad" as const },
  { key: "Licencia", tipo: "cad" as const },
  { key: "Carnet de sanidad", tipo: "emi" as const },
  { key: "Antecedentes penales", tipo: "emi" as const },
  { key: "Antecedentes policiales", tipo: "emi" as const },
];

const makeDefaultDocRows = (): DocRow[] =>
  DOC_DEFS.map((d) => ({
    key: d.key,
    label: d.key,
    tipo: d.tipo,
    fecha: null,
    attachments: [],
    editing: false,
    justUpdated: false,
    file: null,
  }));

// ---- Mini componente: Tarjeta documento ----
interface DocCardProps {
  title: string;
  dateLabel: string;
  dateValue: Date | null;
  onDateChange: (date: Date | null) => void;
  minDate?: Date;
  maxDate?: Date;
  file: File | null;
  onFileChange: (file: File | null) => void;
  attachments?: Attach[];
}

// strings del DatePicker en español
const datePickerStringsEs = {
  months: [
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
  ],
  shortMonths: [
    "ene",
    "feb",
    "mar",
    "abr",
    "may",
    "jun",
    "jul",
    "ago",
    "sep",
    "oct",
    "nov",
    "dic",
  ],
  days: [
    "domingo",
    "lunes",
    "martes",
    "miércoles",
    "jueves",
    "viernes",
    "sábado",
  ],
  shortDays: ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"],
  goToToday: "Ir a hoy",
  calendarDayFormat: "dddd D",
  monthPickerHeaderAriaLabel: "{0}, elija un mes",
  yearPickerHeaderAriaLabel: "{0}, elija un año",
};

// formato de fecha dd/mm/aaaa
const formatDateEs = (date?: Date) =>
  date ? date.toLocaleDateString("es-ES") : "";

const DocCard: React.FC<DocCardProps> = ({
  title,
  dateLabel,
  dateValue,
  onDateChange,
  minDate,
  maxDate,
  file,
  onFileChange,
  attachments,
}) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const datePickerHostRef = React.useRef<HTMLDivElement>(null);
  const shouldRestoreDateFocusRef = React.useRef(false);
  const scrollSnapshotRef = React.useRef<{
    windowX: number;
    windowY: number;
    containers: Array<{
      element: HTMLElement;
      top: number;
      left: number;
    }>;
  } | null>(null);

  const ocultarArchivo = title === "DNI" || title === "Licencia";

  const captureScrollSnapshot = React.useCallback((): void => {
    const host = datePickerHostRef.current;
    const containers: Array<{
      element: HTMLElement;
      top: number;
      left: number;
    }> = [];

    let current = host?.parentElement ?? null;
    while (current) {
      const style = window.getComputedStyle(current);
      const overflowY = style.overflowY || "";
      const overflowX = style.overflowX || "";
      const isScrollable =
        /(auto|scroll|overlay)/.test(`${overflowY} ${overflowX}`) &&
        (current.scrollHeight > current.clientHeight ||
          current.scrollWidth > current.clientWidth);

      if (isScrollable) {
        containers.push({
          element: current,
          top: current.scrollTop,
          left: current.scrollLeft,
        });
      }

      current = current.parentElement;
    }

    scrollSnapshotRef.current = {
      windowX: window.scrollX,
      windowY: window.scrollY,
      containers,
    };
  }, []);

  const restoreDateInputFocus = React.useCallback((): void => {
    if (!shouldRestoreDateFocusRef.current) return;

    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        const snapshot = scrollSnapshotRef.current;
        if (snapshot) {
          for (let i = 0; i < snapshot.containers.length; i++) {
            const container = snapshot.containers[i];
            container.element.scrollTop = container.top;
            container.element.scrollLeft = container.left;
          }

          window.scrollTo(snapshot.windowX, snapshot.windowY);
        }

        const input = datePickerHostRef.current?.querySelector("input");
        shouldRestoreDateFocusRef.current = false;
        scrollSnapshotRef.current = null;
        if (!input) return;

        try {
          input.focus({ preventScroll: true });
        } catch {
          input.focus();
        }
      });
    });
  }, []);

  return (
    <Stack
      tokens={{ childrenGap: 8 }}
      styles={{
        root: {
          border: `1px solid ${BRAND.border}`,
          borderRadius: 22,
          padding: 14,
          boxShadow: "0 16px 30px rgba(0,87,166,.08)",
          background: "linear-gradient(180deg, rgba(255,255,255,.98) 0%, #f6fbff 100%)",
          minWidth: 240,
          maxWidth: 260,
        },
      }}
    >
      <Label
        styles={{
          root: {
            fontWeight: 600,
            color: theme.palette.themePrimary,
          },
        }}
      >
        {title}
      </Label>

      <div ref={datePickerHostRef}>
        <DatePicker
          label={dateLabel}
          value={dateValue || undefined}
          onSelectDate={(d) => {
            shouldRestoreDateFocusRef.current = true;
            captureScrollSnapshot();
            onDateChange(d ?? null);
          }}
          onAfterMenuDismiss={restoreDateInputFocus}
          minDate={minDate}
          maxDate={maxDate}
          firstDayOfWeek={DayOfWeek.Monday}
          placeholder="dd/mm/aaaa"
          ariaLabel={dateLabel}
          strings={datePickerStringsEs}
          formatDate={formatDateEs}
          styles={roundedDatePicker}
          // ✅ FIX: evita salto de scroll/foco al top cuando está dentro de un Modal/Dialog
          calloutProps={{
            doNotLayer: true,
            setInitialFocus: false,
          }}
        />
      </div>

      {!ocultarArchivo && (
        <div>
          <Label>Adjuntar archivo</Label>
          <input
            ref={fileInputRef}
            type="file"
            style={{ display: "none" }}
            onChange={(e) =>
              onFileChange(
                e.target.files && e.target.files.length
                  ? e.target.files[0]
                  : null
              )
            }
          />
          <DefaultButton
            text="Adjuntar archivo"
            iconProps={{ iconName: "Upload" }}
            onClick={() => fileInputRef.current?.click()}
            styles={secondaryButtonStyles}
          />
          <div style={{ fontSize: 12, marginTop: 4 }}>
            {file ? file.name : "-"}
          </div>
        </div>
      )}

      {attachments && attachments.length > 0 && (
        <div style={{ marginTop: 8 }}>
          <Label>Archivos actuales</Label>
          <Stack tokens={{ childrenGap: 4 }}>
            {attachments.map((a) => (
              <a
                key={a.href}
                href={a.href}
                target="_blank"
                rel="noopener noreferrer"
              >
                {a.name}
              </a>
            ))}
          </Stack>
        </div>
      )}
    </Stack>
  );
};

// ===== Helpers de fechas =====
const today0 = () => {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
};

const addMonthsSafe = (d: Date, months: number) => {
  const x = new Date(d.getTime());
  const day = x.getDate();
  x.setMonth(x.getMonth() + months);
  if (x.getDate() < day) x.setDate(0);
  return x;
};

const isNotOlderThanMonths = (d: Date | null, months: number) => {
  if (!d) return true;
  return addMonthsSafe(d, months) >= today0();
};

const isNotExpired = (d?: Date): boolean => {
  if (!d) return true;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()) >= today0();
};

const cutoffSinceMonths = (months: number) => addMonthsSafe(today0(), -months);

// ===========================================================
// REGISTRO DE PERSONAL
// ===========================================================
const RegistroPersonal: React.FC<IRegistroPersonalProps> = ({
  sp,
  siteUrl,
  filtrarPorProveedor,
  borrar,
  bloquearEmpresa,
}) => {
  const [modo, setModo] = React.useState<Modo>("Ingresar");

  // Proveedor seleccionado / detectado
  const [proveedorTitleOculto, setProveedorTitleOculto] = React.useState("");
  const [proveedorId, setProveedorId] = React.useState<number | null>(null);

  // Dropdown proveedores (modo editable)
  const [proveedoresOptions, setProveedoresOptions] = React.useState<
    IDropdownOption[]
  >([]);
  const proveedoresByIdRef = React.useRef<Map<number, string>>(new Map());

  // ✅ fuerza remount del form (limpia controles con estado interno, ej: Dropdown/File)
  const [formKey, setFormKey] = React.useState(0);

  // ✅ ancla superior para volver arriba al limpiar
  const topRef = React.useRef<HTMLDivElement | null>(null);
  const scrollToTop = () => {
    if (topRef.current) {
      topRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
      return;
    }
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const [form, setForm] = React.useState<PersonaForm>({
    Documento: "",
    Nombre: "",
    ApellidoPaterno: "",
    ApellidoMaterno: "",
    TipoDocumento: undefined,
    Puesto: undefined,
    Especificar: "",
    Licencia: "",
    Categoria: undefined,
    CorreosNotificacion: "",
  });

  const isDarDeBaja = modo === "Dar de baja";

  // ✅ Touch: cuando el usuario edita cualquier cosa, limpiamos el error para re-habilitar Guardar
  const [guardando, setGuardando] = React.useState(false);
  const [mensaje, setMensaje] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);

  const touch = () => {
    setError(null);
    // setMensaje(null);
  };

  const setDateAndTouch =
    (setter: (d: Date | null) => void) => (d: Date | null) => {
      touch();
      setter(d);
    };

  const setFileAndTouch =
    (setter: (f: File | null) => void) => (f: File | null) => {
      touch();
      setter(f);
    };

  // -------- Sección 3: Documentación (para Ingresar) --------
  const [dniCaducidad, setDniCaducidad] = React.useState<Date | null>(null);
  const [dniFile, setDniFile] = React.useState<File | null>(null);
  const [licCaducidad, setLicCaducidad] = React.useState<Date | null>(null);
  const [licFile, setLicFile] = React.useState<File | null>(null);
  const [carnetEmision, setCarnetEmision] = React.useState<Date | null>(null);
  const [carnetFile, setCarnetFile] = React.useState<File | null>(null);
  const [penalesEmision, setPenalesEmision] = React.useState<Date | null>(null);
  const [penalesFile, setPenalesFile] = React.useState<File | null>(null);
  const [policialesEmision, setPolicialesEmision] = React.useState<Date | null>(
    null
  );
  const [policialesFile, setPolicialesFile] = React.useState<File | null>(null);

  // ---- Estado de documentación (Modificar) ----
  const [docRows, setDocRows] = React.useState<DocRow[]>(makeDefaultDocRows());

  // confirmación de baja
  const [showConfirmBaja, setShowConfirmBaja] = React.useState(false);
  const [motivoBaja, setMotivoBaja] = React.useState("");

  // ===== visibilidad dinámica por Puesto =====
  const puestoNorm = (form.Puesto || "").toLowerCase().trim();
  const showEspecificar = React.useMemo(() => puestoNorm === "otro", [puestoNorm]);
  const showLicenciaCat = React.useMemo(() => puestoNorm === "conductor", [puestoNorm]);

  // ---- Validación de antigüedad ----
  const errorDocs = React.useMemo(() => {
    const fmt = (d: Date) => d.toLocaleDateString();

    if (modo === "Ingresar") {
      if (!isNotExpired(dniCaducidad || undefined)) {
        return 'La fecha de caducidad del "DNI" no puede estar vencida.';
      }
      if (showLicenciaCat && !isNotExpired(licCaducidad || undefined)) {
        return 'La fecha de caducidad de la "Licencia" no puede estar vencida.';
      }
      if (!isNotOlderThanMonths(carnetEmision, 6)) {
        return `La fecha de emisión del "Carnet de sanidad" no puede ser anterior a ${fmt(
          cutoffSinceMonths(6)
        )}.`;
      }
      if (!isNotOlderThanMonths(penalesEmision, 12)) {
        return `La fecha de emisión de "Antecedentes penales" no puede ser anterior a ${fmt(
          cutoffSinceMonths(12)
        )}.`;
      }
      if (!isNotOlderThanMonths(policialesEmision, 12)) {
        return `La fecha de emisión de "Antecedentes policiales" no puede ser anterior a ${fmt(
          cutoffSinceMonths(12)
        )}.`;
      }
      return null;
    }

    if (modo === "Modificar") {
      const getFechaByLabel = (lbl: string): Date | null => {
        for (let i = 0; i < docRows.length; i++) {
          if (docRows[i].label === lbl) return docRows[i].fecha || null;
        }
        return null;
      };

      const dni = getFechaByLabel("DNI");
      if (!isNotExpired(dni || undefined)) {
        return 'La fecha de caducidad del "DNI" no puede estar vencida.';
      }
      const lic = getFechaByLabel("Licencia");
      if (!isNotExpired(lic || undefined)) {
        return 'La fecha de caducidad de la "Licencia" no puede estar vencida.';
      }
      const cSan = getFechaByLabel("Carnet de sanidad");
      if (!isNotOlderThanMonths(cSan, 6)) {
        return `La fecha de emisión del "Carnet de sanidad" no puede ser anterior a ${fmt(
          cutoffSinceMonths(6)
        )}.`;
      }
      const pen = getFechaByLabel("Antecedentes penales");
      if (!isNotOlderThanMonths(pen, 12)) {
        return `La fecha de emisión de "Antecedentes penales" no puede ser anterior a ${fmt(
          cutoffSinceMonths(12)
        )}.`;
      }
      const pol = getFechaByLabel("Antecedentes policiales");
      if (!isNotOlderThanMonths(pol, 12)) {
        return `La fecha de emisión de "Antecedentes policiales" no puede ser anterior a ${fmt(
          cutoffSinceMonths(12)
        )}.`;
      }
      return null;
    }

    return null;
  }, [modo, dniCaducidad, licCaducidad, showLicenciaCat, carnetEmision, penalesEmision, policialesEmision, docRows]);

  // ✅ NUEVO: foco automático solo al error principal (no a validaciones en vivo)
  const errorRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    if (!error) return;

    const target = errorRef.current;
    if (!target) return;

    requestAnimationFrame(() => {
      try {
        target.focus();
        target.scrollIntoView({ behavior: "smooth", block: "start" });
      } catch {
        // ignore
      }
    });
  }, [error]);

  // ===== bloqueo de Datos laborales según documento =====
  const documentoLengthRequerido = React.useMemo(
    () => getDocumentoLengthRequerido(form.TipoDocumento),
    [form.TipoDocumento]
  );

  const documentoAyuda = React.useMemo(() => {
    if (!form.TipoDocumento || !documentoLengthRequerido) {
      return "Selecciona un tipo de documento para ver la cantidad de caracteres requerida.";
    }

    return `${form.TipoDocumento}: debe tener ${documentoLengthRequerido} caracteres.`;
  }, [form.TipoDocumento, documentoLengthRequerido]);

  const documentoValido = React.useMemo(() => {
    const len = (form.Documento || "").trim().length;

    if (!documentoLengthRequerido) return false;
    return len === documentoLengthRequerido;
  }, [documentoLengthRequerido, form.Documento]);

  const documentoErrorMessage = React.useMemo(() => {
    if (!form.Documento?.trim() || !documentoLengthRequerido || documentoValido) {
      return undefined;
    }

    return `${form.TipoDocumento}: debe tener ${documentoLengthRequerido} caracteres.`;
  }, [documentoLengthRequerido, documentoValido, form.Documento, form.TipoDocumento]);

  const laboralBloqueado = React.useMemo(() => !documentoValido, [documentoValido]);

  // =======================
  // Proveedores: cargar opciones SIEMPRE (para el modo editable)
  // =======================
  React.useEffect(() => {
    let cancelado = false;

    const cargarProveedores = async () => {
      try {
        const items = await sp.web.lists
          .getByTitle(LST_PROVEEDORES)
          .items.select("Id", "Title")
          .orderBy("Title", true)
          .top(5000)();

        if (cancelado) return;

        const map = new Map<number, string>();
        const opts: IDropdownOption[] = (items as any[]).map((it) => {
          const id = Number(it.Id);
          const title = String(it.Title || "");
          map.set(id, title);
          return { key: id, text: title };
        });

        proveedoresByIdRef.current = map;
        setProveedoresOptions(opts);

        if (!bloquearEmpresa && proveedorId && !proveedorTitleOculto) {
          const t = map.get(proveedorId) || "";
          if (t) setProveedorTitleOculto(t);
        }
      } catch {
        if (!cancelado) {
          proveedoresByIdRef.current = new Map();
          setProveedoresOptions([]);
        }
      }
    };

    cargarProveedores().catch(() => {});
    return () => {
      cancelado = true;
    };
  }, [sp, bloquearEmpresa, proveedorId, proveedorTitleOculto]);

  // =======================
  // Meta lookup Proveedor (para armar payload correcto)
  // =======================
  type LookupMeta = {
    InternalName: string;
    TypeAsString: string;
    AllowMultipleValues?: boolean;
  };
  const [provFieldMeta, setProvFieldMeta] = React.useState<LookupMeta | null>(null);

  const buildProveedorPayload = () => {
    if (!proveedorId || !provFieldMeta) return {};
    const key = `${provFieldMeta.InternalName}Id`;
    const tas = (provFieldMeta.TypeAsString || "").toLowerCase();
    const isMulti =
      provFieldMeta.AllowMultipleValues === true || tas.indexOf("multi") !== -1;
    return isMulti ? { [key]: { results: [proveedorId] } } : { [key]: proveedorId };
  };

  // =======================
  // Si bloquearEmpresa=true => autodetectar proveedor por usuario
  // Si bloquearEmpresa=false => NO pisar selección
  // =======================
  React.useEffect(() => {
    let cancelado = false;

    const cargarMetaYProveedorUsuario = async () => {
      try {
        const f = await sp.web.lists
          .getByTitle(LST_PERSONAS)
          .fields.getByInternalNameOrTitle("Proveedor")
          .select("InternalName", "TypeAsString", "AllowMultipleValues")();

        if (!cancelado)
          setProvFieldMeta({
            InternalName: f.InternalName,
            TypeAsString: f.TypeAsString,
            AllowMultipleValues: (f as any).AllowMultipleValues,
          });

        if (!bloquearEmpresa) return;

        const me = await sp.web.currentUser();
        let items = await sp.web.lists
          .getByTitle(LST_PROVEEDORES)
          .items.select("Id", "Title", "Usuarios/Id")
          .expand("Usuarios")
          .filter(`Usuarios/Id eq ${me.Id}`)
          .top(1)();

        if (!items?.length) {
          items = await sp.web.lists
            .getByTitle(LST_PROVEEDORES)
            .items.select("Id", "Title", "UsuariosId")
            .filter(`UsuariosId eq ${me.Id}`)
            .top(1)();
        }

        if (!cancelado && items?.[0]) {
          setProveedorTitleOculto(items[0].Title);
          setProveedorId(items[0].Id);
        }
      } catch {
        // ignoramos
      }
    };

    cargarMetaYProveedorUsuario().catch(() => {});
    return () => {
      cancelado = true;
    };
  }, [sp, bloquearEmpresa]);

  const onChange = (field: keyof PersonaForm, value?: string) => {
    touch();
    setForm((prev) => ({ ...prev, [field]: value ?? "" }));
  };

  // --- Filtro local de la grilla de personas ---
  type PersonaItem = {
    Id: number;
    Title: string;
    Nombre?: string;
    Apellido_x0020_paterno?: string;
    Apellido_x0020_materno?: string;
    tipodocumento?: string;
    puesto?: string;
    otro?: string;
    Licencia?: string;
    Categoria?: string;
    ProveedorId?: number;
    correosnotificacion?: string;
  };

  const columns: IColumn[] = [
    { key: "doc", name: "Documento", fieldName: "Title", minWidth: 120, isResizable: true },
    { key: "nom", name: "Nombre", fieldName: "Nombre", minWidth: 120, isResizable: true },
    {
      key: "ap",
      name: "Ap. paterno",
      fieldName: "Apellido_x0020_paterno",
      minWidth: 120,
      isResizable: true,
    },
    {
      key: "am",
      name: "Ap. materno",
      fieldName: "Apellido_x0020_materno",
      minWidth: 120,
      isResizable: true,
    },
    { key: "pto", name: "Puesto", fieldName: "puesto", minWidth: 120, isResizable: true },
    { key: "cat", name: "Categoría", fieldName: "Categoria", minWidth: 90, isResizable: true },
  ];

  const [itemsProveedor, setItemsProveedor] = React.useState<PersonaItem[]>([]);
  const [cargandoGrid, setCargandoGrid] = React.useState(false);

  const [queryGrid, setQueryGrid] = React.useState("");
  const [gridPage, setGridPage] = React.useState(1);
  const itemsProveedorFiltrados = React.useMemo(() => {
    const q = queryGrid.trim().toLowerCase();
    if (!q) return itemsProveedor;
    return itemsProveedor.filter((it) => {
      const txt = [
        it.Title,
        it.Nombre,
        it.Apellido_x0020_paterno,
        it.Apellido_x0020_materno,
        it.puesto,
        it.Categoria,
        it.tipodocumento,
        it.otro,
        it.Licencia,
      ]
        .map((v) => (v ?? "").toString().toLowerCase())
        .join(" ");
      return txt.indexOf(q) !== -1;
    });
  }, [itemsProveedor, queryGrid]);
  const totalGridPages = React.useMemo(
    () => Math.max(1, Math.ceil(itemsProveedorFiltrados.length / GRID_PAGE_SIZE)),
    [itemsProveedorFiltrados.length]
  );
  const itemsProveedorPagina = React.useMemo(() => {
    const start = (gridPage - 1) * GRID_PAGE_SIZE;
    return itemsProveedorFiltrados.slice(start, start + GRID_PAGE_SIZE);
  }, [gridPage, itemsProveedorFiltrados]);

  React.useEffect(() => {
    setGridPage(1);
  }, [queryGrid, itemsProveedor]);

  React.useEffect(() => {
    setGridPage((prev) => Math.min(prev, totalGridPages));
  }, [totalGridPages]);

  const selectionRef = React.useRef<Selection>();
  if (!selectionRef.current) {
    selectionRef.current = new Selection({
      onSelectionChanged: () => {
        const sel = selectionRef.current!.getSelection() as PersonaItem[];
        if (sel && sel.length) {
          loadFromGridItem(sel[0]).catch((e) =>
            console.warn("selection -> loadFromGridItem:", e)
          );
        }
      },
    });
  }

  React.useEffect(() => {
    const visible = modo === "Modificar" || modo === "Dar de baja";
    if (!visible) {
      setItemsProveedor([]);
      return;
    }

    let cancelado = false;

    const cargarGrid = async () => {
      setCargandoGrid(true);
      try {
        let query = sp.web.lists
          .getByTitle(LST_PERSONAS)
          .items.select(
            "Id",
            "Title",
            "Nombre",
            "Apellido_x0020_paterno",
            "Apellido_x0020_materno",
            "tipodocumento",
            "puesto",
            "otro",
            "Licencia",
            "Categoria",
            "ProveedorId",
            "correosnotificacion",
            "activo"
          )
          .orderBy("Id", false)
          .top(5000);

        let filter = "activo eq 1";

        if (filtrarPorProveedor && proveedorId) {
          filter += ` and ProveedorId eq ${proveedorId}`;
        }

        query = query.filter(filter);

        const items = await query();

        const itemsFiltrados =
          filtrarPorProveedor && proveedorId
            ? (items as any[]).filter((it) => it.ProveedorId === proveedorId)
            : (items as any[]);

        if (!cancelado) setItemsProveedor(itemsFiltrados);
      } catch {
        if (!cancelado) setItemsProveedor([]);
      } finally {
        if (!cancelado) setCargandoGrid(false);
      }
    };

    cargarGrid().catch(() => {});
    return () => {
      cancelado = true;
    };
  }, [modo, sp, filtrarPorProveedor, proveedorId]);

  React.useEffect(() => {
    if (modo === "Modificar") {
      if (form.Documento?.trim()) {
        loadDocumentacionByTitle(form.Documento).catch((e) =>
          console.warn("auto load docs by Documento:", e)
        );
      } else {
        setDocRows(makeDefaultDocRows());
      }
    }
  }, [modo, form.Documento]);

  // ✅ limpiar configurable + remount + scroll arriba
  const limpiar = (opts?: { keepMessages?: boolean }) => {
    const keepMessages = opts?.keepMessages === true;

    setForm({
      Documento: "",
      Nombre: "",
      ApellidoPaterno: "",
      ApellidoMaterno: "",
      TipoDocumento: undefined,
      Puesto: undefined,
      Especificar: "",
      Licencia: "",
      Categoria: undefined,
      CorreosNotificacion: "",
    });

    setDniCaducidad(null);
    setDniFile(null);
    setLicCaducidad(null);
    setLicFile(null);
    setCarnetEmision(null);
    setCarnetFile(null);
    setPenalesEmision(null);
    setPenalesFile(null);
    setPolicialesEmision(null);
    setPolicialesFile(null);

    setDocRows(makeDefaultDocRows());

    setShowConfirmBaja(false);
    setMotivoBaja("");

    if (!keepMessages) {
      setMensaje(null);
      setError(null);
    } else {
      setError(null);
    }

    setFormKey((k) => k + 1);

    try {
      selectionRef.current?.setAllSelected(false);
    } catch {
      // nada
    }

    requestAnimationFrame(() => scrollToTop());
  };

  const findPersonaByDocumento = async (doc: string) => {
    const items = await sp.web.lists
      .getByTitle(LST_PERSONAS)
      .items.select("Id", "Title")
      .filter(`Title eq '${esc(doc)}'`)();
    return items[0];
  };

  const crearEnPersonas = async () =>
    sp.web.lists.getByTitle(LST_PERSONAS).items.add({
      Title: form.Documento,
      Nombre: form.Nombre,
      Apellido_x0020_paterno: form.ApellidoPaterno,
      Apellido_x0020_materno: form.ApellidoMaterno,
      tipodocumento: form.TipoDocumento,
      puesto: form.Puesto,
      otro: form.Especificar,
      Licencia: form.Licencia,
      Categoria: form.Categoria,
      correosnotificacion: stripHtml(form.CorreosNotificacion),
      ...buildProveedorPayload(),
    });

  const actualizarEnPersonas = async (id: number) =>
    sp.web.lists
      .getByTitle(LST_PERSONAS)
      .items.getById(id)
      .update({
        Nombre: form.Nombre,
        Apellido_x0020_paterno: form.ApellidoPaterno,
        Apellido_x0020_materno: form.ApellidoMaterno,
        tipodocumento: form.TipoDocumento,
        puesto: form.Puesto,
        otro: form.Especificar,
        Licencia: form.Licencia,
        Categoria: form.Categoria,
        correosnotificacion: stripHtml(form.CorreosNotificacion),
        ...buildProveedorPayload(),
      });

  const eliminarEnPersonas = async (id: number) =>
    sp.web.lists.getByTitle(LST_PERSONAS).items.getById(id).delete();

  const addDocItem = async (label: string, fields: DocFields): Promise<number> => {
    const payload: any = { Title: form.Documento, Documento: label };
    if (fields.Caducidad !== undefined) payload.Caducidad = fields.Caducidad;
    if (fields.Emision !== undefined) payload.Emision = fields.Emision;

    const add = await sp.web.lists.getByTitle(LST_DOCS).items.add(payload);
    const idFromData = Number((add as any)?.data?.Id ?? (add as any)?.data?.ID);
    if (idFromData && !isNaN(idFromData)) return idFromData;

    const found = await sp.web.lists
      .getByTitle(LST_DOCS)
      .items.select("Id")
      .filter(`Title eq '${esc(form.Documento)}' and Documento eq '${esc(label)}'`)
      .orderBy("Id", false)
      .top(1)();
    const id = Number(found?.[0]?.Id);
    if (!id || isNaN(id))
      throw new Error("No se pudo obtener el Id del item de Documentación.");
    return id;
  };

  const updateDocItem = async (id: number, fields: DocFields) => {
    const payload: any = {};
    if (fields.Caducidad !== undefined) payload.Caducidad = fields.Caducidad;
    if (fields.Emision !== undefined) payload.Emision = fields.Emision;
    await sp.web.lists.getByTitle(LST_DOCS).items.getById(id).update(payload);
  };

  const deleteAllDocsByTitle = async (docTitle: string) => {
    const items = await sp.web.lists
      .getByTitle(LST_DOCS)
      .items.select("Id")
      .filter(`Title eq '${esc(docTitle)}'`)();
    for (const it of items)
      await sp.web.lists.getByTitle(LST_DOCS).items.getById(it.Id).delete();
  };

  const getDocItemByLabel = async (docTitle: string, label: string) => {
    const items = await sp.web.lists
      .getByTitle(LST_DOCS)
      .items.select("Id", "Title", "Documento")
      .filter(`Title eq '${esc(docTitle)}' and Documento eq '${esc(label)}'`)();
    return items[0];
  };

  const attachFile = async (itemId: number, file: File) => {
    const item = sp.web.lists.getByTitle(LST_DOCS).items.getById(itemId);
    try {
      const current = await item.attachmentFiles();
      if (current && current.length) {
        for (let i = 0; i < current.length; i++) {
          try {
            await item.attachmentFiles.getByName(current[i].FileName).delete();
          } catch {
            // ignoramos
          }
        }
      }
    } catch {
      // ignoramos
    }
    await item.attachmentFiles.add(file.name, file);
  };

  const upsertDocRow = async (
    label: string,
    fields: DocFields,
    file?: File | null
  ) => {
    if (!form.Documento?.trim())
      throw new Error("Documento es obligatorio para Documentación.");
    const existing = await getDocItemByLabel(form.Documento, label);
    let id: number;
    if (existing?.Id) {
      await updateDocItem(Number(existing.Id), fields);
      id = Number(existing.Id);
    } else {
      id = await addDocItem(label, fields);
    }
    if (!id || isNaN(id)) throw new Error("No se pudo obtener Id de Documentación.");
    if (file) await attachFile(id, file);
  };

  function toDate(iso?: string | null): Date | null {
    if (!iso) return null;
    const d = new Date(iso);
    return isNaN(d.getTime()) ? null : d;
  }

  async function loadDocumentacionByTitle(docTitle: string): Promise<void> {
    const rows = await sp.web.lists
      .getByTitle(LST_DOCS)
      .items.select("Id", "Title", "Documento", "Caducidad", "Emision")
      .filter(`Title eq '${esc(docTitle)}'`)
      .top(5000)();

    const map = new Map<string, any>(rows.map((r) => [String(r.Documento), r]));

    setDniCaducidad(toDate(map.get("DNI")?.Caducidad ?? null));
    setLicCaducidad(toDate(map.get("Licencia")?.Caducidad ?? null));
    setCarnetEmision(toDate(map.get("Carnet de sanidad")?.Emision ?? null));
    setPenalesEmision(toDate(map.get("Antecedentes penales")?.Emision ?? null));
    setPolicialesEmision(toDate(map.get("Antecedentes policiales")?.Emision ?? null));

    const defs = [
      { key: "DNI", tipo: "cad" as const, fechaRaw: map.get("DNI")?.Caducidad ?? null },
      { key: "Licencia", tipo: "cad" as const, fechaRaw: map.get("Licencia")?.Caducidad ?? null },
      {
        key: "Carnet de sanidad",
        tipo: "emi" as const,
        fechaRaw: map.get("Carnet de sanidad")?.Emision ?? null,
      },
      {
        key: "Antecedentes penales",
        tipo: "emi" as const,
        fechaRaw: map.get("Antecedentes penales")?.Emision ?? null,
      },
      {
        key: "Antecedentes policiales",
        tipo: "emi" as const,
        fechaRaw: map.get("Antecedentes policiales")?.Emision ?? null,
      },
    ];

    const withAtts: DocRow[] = [];
    for (let i = 0; i < defs.length; i++) {
      const d = defs[i];
      const r = map.get(d.key);
      let attachments: Attach[] = [];
      if (r?.Id) {
        try {
          const atts = await sp.web.lists
            .getByTitle(LST_DOCS)
            .items.getById(r.Id)
            .attachmentFiles();
          attachments = (atts || []).map((a: any) => ({
            name: a.FileName,
            href: toAbs(siteUrl, a.ServerRelativeUrl),
          }));
        } catch {
          // ignoramos
        }
      }
      withAtts.push({
        key: d.key,
        label: d.key,
        tipo: d.tipo,
        fecha: toDate(d.fechaRaw),
        attachments,
      });
    }

    setDocRows(withAtts);

    setDniFile(null);
    setLicFile(null);
    setCarnetFile(null);
    setPenalesFile(null);
    setPolicialesFile(null);
  }

  async function loadFromGridItem(it: any) {
    touch();

    if (!bloquearEmpresa) {
      const pid = it.ProveedorId ? Number(it.ProveedorId) : null;
      setProveedorId(pid);
      if (pid) {
        const t = proveedoresByIdRef.current.get(pid) || "";
        setProveedorTitleOculto(t);
      } else {
        setProveedorTitleOculto("");
      }
    }

    setForm({
      Documento: it.Title ?? "",
      Nombre: it.Nombre ?? "",
      ApellidoPaterno: it.Apellido_x0020_paterno ?? "",
      ApellidoMaterno: it.Apellido_x0020_materno ?? "",
      TipoDocumento: it.tipodocumento ?? undefined,
      Puesto: it.puesto ?? undefined,
      Especificar: it.otro ?? "",
      Licencia: it.Licencia ?? "",
      Categoria: it.Categoria ?? undefined,
      CorreosNotificacion: stripHtml(it.correosnotificacion),
    });

    if (it.Title) {
      try {
        await loadDocumentacionByTitle(it.Title);
      } catch (e) {
        console.warn(e);
      }
    }
  }

  const onCancelar = () => limpiar();

  const modoOptions: Array<{ key: Modo; text: string; iconName: string }> = [
    { key: "Ingresar", text: "Ingresar", iconName: "Add" },
    { key: "Modificar", text: "Modificar", iconName: "Edit" },
    { key: "Dar de baja", text: "Dar de baja", iconName: "Delete" },
  ];

  // ======= validación de docs obligatorios para Ingresar =======
  const docsObligIngresar =
    dniCaducidad !== null &&
    carnetEmision !== null &&
    carnetFile !== null &&
    penalesEmision !== null &&
    penalesFile !== null &&
    policialesEmision !== null &&
    policialesFile !== null &&
    (!showLicenciaCat || licCaducidad !== null);

  const requiredFieldsPendientesIngresar = React.useMemo(() => {
    if (modo !== "Ingresar") return [];

    const pendientes: string[] = [];

    if (!bloquearEmpresa && !proveedorId) pendientes.push("Empresa");
    if (!form.Nombre?.trim()) pendientes.push("Nombre");
    if (!form.Documento?.trim()) pendientes.push("Documento");
    if (dniCaducidad === null) pendientes.push("Fecha de caducidad del DNI");
    if (showLicenciaCat && licCaducidad === null) {
      pendientes.push("Fecha de caducidad de la Licencia");
    }
    if (carnetEmision === null) pendientes.push("Fecha de emision del Carnet de sanidad");
    if (carnetFile === null) pendientes.push("Archivo del Carnet de sanidad");
    if (penalesEmision === null) {
      pendientes.push("Fecha de emision de Antecedentes penales");
    }
    if (penalesFile === null) pendientes.push("Archivo de Antecedentes penales");
    if (policialesEmision === null) {
      pendientes.push("Fecha de emision de Antecedentes policiales");
    }
    if (policialesFile === null) pendientes.push("Archivo de Antecedentes policiales");

    return pendientes;
  }, [
    bloquearEmpresa,
    proveedorId,
    form.Nombre,
    form.Documento,
    modo,
    dniCaducidad,
    showLicenciaCat,
    licCaducidad,
    carnetEmision,
    carnetFile,
    penalesEmision,
    penalesFile,
    policialesEmision,
    policialesFile,
  ]);

  const puedeGuardar =
    !guardando &&
    !errorDocs &&
    (modo !== "Ingresar" || requiredFieldsPendientesIngresar.length === 0);

  const fechaHoy = today0();

  // ✅ MODIFICADO: mensajes de error “de negocio” sin usar startsWith/includes (compat TS lib vieja)
  const getFriendlyError = (e: any): string => {
    const raw =
      e?.data?.error?.message?.value ||
      e?.data?.error?.message ||
      e?.odata?.error?.message?.value ||
      e?.message ||
      e?.statusText ||
      String(e);

    const rawStr = String(raw || "").trim();

    const tryExtractFromEmbeddedJson = (s: string): string | null => {
      const arrowIdx = s.indexOf("=>");
      const candidate = (arrowIdx >= 0 ? s.slice(arrowIdx + 2) : s).trim();

      if (candidate && candidate.charAt(0) === "{") {
        try {
          const obj = JSON.parse(candidate);
          const val =
            (obj &&
              obj["odata.error"] &&
              obj["odata.error"].message &&
              obj["odata.error"].message.value) ||
            (obj &&
              obj.odata &&
              obj.odata.error &&
              obj.odata.error.message &&
              obj.odata.error.message.value) ||
            (obj && obj.error && obj.error.message && obj.error.message.value) ||
            (obj && obj.error && obj.error.message);
          if (val) return String(val);
        } catch {
          // ignore
        }
      }

      const m = s.match(
        /"message"\s*:\s*\{\s*"lang"\s*:\s*"[^"]*"\s*,\s*"value"\s*:\s*"((?:\\.|[^"\\])*)"\s*\}/i
      );
      if (m && m[1]) {
        try {
          return JSON.parse(`"${m[1]}"`);
        } catch {
          return m[1];
        }
      }
      return null;
    };

    const stripTechPrefix = (s: string) =>
      s
        .replace(/^Error:\s*/i, "")
        .replace(/^Error making HttpClient request.*?=>\s*/i, "")
        .replace(/\s+/g, " ")
        .trim();

    const extracted = tryExtractFromEmbeddedJson(rawStr);
    const msg = stripTechPrefix(extracted ?? rawStr);

    const lower = msg.toLowerCase();

    if (
      lower.indexOf("spduplicatevaluesfoundexception") !== -1 ||
      lower.indexOf("valores duplicados") !== -1 ||
      lower.indexOf("duplicate") !== -1
    ) {
      const doc = (form.Documento || "").trim();
      return doc
        ? `Ya existe un registro con el documento ${doc}. No se puede guardar duplicado.`
        : "Ya existe un registro con ese documento. No se puede guardar duplicado.";
    }

    const idxElElemento = lower.indexOf("el elemento");
    const msg2 = idxElElemento >= 0 ? msg.slice(idxElElemento).trim() : msg;

    if (!msg2 || msg2.toLowerCase() === "error") {
      return "No se pudo guardar. Revisá los datos e intentá nuevamente.";
    }

    return `No se pudo guardar: ${msg2}`;
  };

  const guardarInterno = async () => {
    setGuardando(true);
    try {
      if (modo === "Ingresar") {
        await crearEnPersonas();

        await upsertDocRow(
          "DNI",
          { Caducidad: dateToISO(dniCaducidad) },
          dniFile ?? undefined
        );

        if (showLicenciaCat) {
          await upsertDocRow(
            "Licencia",
            { Caducidad: dateToISO(licCaducidad) },
            licFile ?? undefined
          );
        }

        await upsertDocRow(
          "Carnet de sanidad",
          { Emision: dateToISO(carnetEmision) },
          carnetFile ?? undefined
        );
        await upsertDocRow(
          "Antecedentes penales",
          { Emision: dateToISO(penalesEmision) },
          penalesFile ?? undefined
        );
        await upsertDocRow(
          "Antecedentes policiales",
          { Emision: dateToISO(policialesEmision) },
          policialesFile ?? undefined
        );

        limpiar({ keepMessages: true });
        setMensaje("Ingresado en Personas y Documentación.");
      }

      if (modo === "Modificar") {
        const persona = await findPersonaByDocumento(form.Documento);
        if (!persona) throw new Error("No existe persona con ese Documento.");
        await actualizarEnPersonas(persona.Id);

        for (let i = 0; i < docRows.length; i++) {
          const r = docRows[i];
          const fields =
            r.tipo === "cad"
              ? { Caducidad: dateToISO(r.fecha) }
              : { Emision: dateToISO(r.fecha) };
          try {
            await upsertDocRow(r.label, fields, r.file || undefined);
          } catch (e) {
            console.warn("upsert", r.label, e);
          }
        }

        await loadDocumentacionByTitle(form.Documento);
        setMensaje("Registro modificado. Documentación actualizada sin borrar adjuntos.");
      }

      if (modo === "Dar de baja") {
        const persona = await findPersonaByDocumento(form.Documento);
        if (!persona) throw new Error("No existe persona con ese Documento.");

        const itemPersona = sp.web.lists.getByTitle(LST_PERSONAS).items.getById(persona.Id);

        if (borrar) {
          await itemPersona.update({ motivobaja: motivoBaja });
          await eliminarEnPersonas(persona.Id);
          await deleteAllDocsByTitle(form.Documento);
          setMensaje("Registro eliminado de Personas y toda la Documentación.");
        } else {
          await itemPersona.update({ activo: false, motivobaja: motivoBaja });
          setMensaje("Registro marcado como inactivo.");
        }
      }

      setDniFile(null);
      setMotivoBaja("");
      setShowConfirmBaja(false);
    } catch (e: any) {
      setError(getFriendlyError(e));
    } finally {
      setGuardando(false);
    }
  };

  // ======= onGuardar =======
  const onGuardar = async () => {
    setMensaje(null);
    setError(null);

    if (!bloquearEmpresa && !proveedorId) {
      setError("Empresa es obligatoria.");
      return;
    }

    if (!form.Documento?.trim()) {
      setError("Documento es obligatorio.");
      return;
    }
    if (documentoErrorMessage) {
      setError(documentoErrorMessage);
      return;
    }
    if (modo !== "Dar de baja" && !form.Nombre?.trim()) {
      setError("Nombre es obligatorio.");
      return;
    }
    if (errorDocs) {
      // Importante: usamos setError para que el foco vaya al banner principal
      setError(errorDocs);
      return;
    }

    if (modo === "Dar de baja" && !showConfirmBaja) {
      setShowConfirmBaja(true);
      return;
    }

    await guardarInterno();
  };

  const onConfirmarBaja = async () => {
    if (!motivoBaja.trim()) return;
    await guardarInterno();
  };

  return (
    <ThemeProvider theme={theme}>
      <Stack
        key={formKey}
        tokens={stackTokens}
        styles={pageShellStyles}
        data-is-scrollable="true"
      >
        {/* ✅ Ancla para volver arriba al limpiar */}
        <div ref={topRef} />

        {/* Barra de modo */}
        <Stack
          tokens={{ childrenGap: 20 }}
          styles={heroPanelStyles}
        >
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 14 }}>
            <div
              style={{
                width: 56,
                height: 56,
                borderRadius: "50%",
                background: "rgba(255,255,255,.16)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                border: "1px solid rgba(255,255,255,.24)",
              }}
            >
              <Icon
                iconName="Contact"
                styles={{ root: { fontSize: 26, color: theme.palette.white } }}
              />
            </div>
            <div>
              <div
                style={{
                  fontSize: 32,
                  fontWeight: 700,
                  lineHeight: 1.05,
                  color: theme.palette.white,
                }}
              >
                Personal
              </div>
            </div>
          </Stack>

          <Stack horizontal wrap tokens={{ childrenGap: 14 }}>
            {modoOptions.map((opt) => (
              <Stack
                key={opt.key}
                styles={modeTileWrapStyles}
              >
                <DefaultButton
                  text={opt.text}
                  iconProps={{ iconName: opt.iconName }}
                  checked={modo === opt.key}
                  styles={modeButtonStyles}
                  onClick={() => {
                    touch();
                    const next = opt.key;
                    setModo(next);
                    if (next === "Ingresar") {
                      limpiar();
                      try {
                        selectionRef.current?.setAllSelected(false);
                      } catch {
                        // nada
                      }
                    }
                  }}
                />
                {modo === opt.key && <span style={modeTileDotStyles} />}
              </Stack>
            ))}
          </Stack>
        </Stack>

        {/* Grilla de personas del proveedor */}
        {(modo === "Modificar" || modo === "Dar de baja") && (
          <Stack
            tokens={{ childrenGap: 8 }}
            styles={sectionCardStyles}
          >
            <Label styles={sectionTitleStyles}>
              Registros del proveedor:{" "}
              {filtrarPorProveedor
                ? proveedorTitleOculto || "(sin proveedor seleccionado)"
                : "Todos"}
            </Label>

            {cargandoGrid ? (
              <Spinner label="Cargando registros..." size={SpinnerSize.small} />
            ) : (
              <>
                <SearchBox
                  placeholder="Filtrar por cualquier campo…"
                  value={queryGrid}
                  onChange={(_, v) => {
                    touch();
                    setQueryGrid(v || "");
                  }}
                  onClear={() => {
                    touch();
                    setQueryGrid("");
                  }}
                  styles={{ root: { maxWidth: 340, marginBottom: 8 } }}
                />

                {itemsProveedorFiltrados.length === 0 ? (
                  <div style={infoBannerStyles.root}>
                    <Icon
                      iconName="Info"
                      styles={{ root: infoBannerStyles.icon }}
                    />
                    <div style={infoBannerStyles.text}>
                      {queryGrid ? "Sin resultados para la búsqueda." : "No hay registros."}
                    </div>
                  </div>
                ) : (
                  <Stack tokens={{ childrenGap: 10 }}>
                    <div style={{ width: "100%", overflowX: "auto" }}>
                      <DetailsList
                        items={itemsProveedorPagina}
                        columns={columns}
                        selectionMode={SelectionMode.single}
                        selection={selectionRef.current}
                        compact
                        styles={{ root: { minWidth: 560 } }}
                        onItemInvoked={(it) =>
                          loadFromGridItem(it as any).catch((e) =>
                            console.warn("onItemInvoked -> loadFromGridItem:", e)
                          )
                        }
                      />
                    </div>

                    <Stack
                      horizontal
                      wrap
                      verticalAlign="center"
                      horizontalAlign="space-between"
                      tokens={{ childrenGap: 10 }}
                    >
                      <div style={{ fontSize: 13, color: BRAND.muted }}>
                        Mostrando {(gridPage - 1) * GRID_PAGE_SIZE + 1}-
                        {Math.min(gridPage * GRID_PAGE_SIZE, itemsProveedorFiltrados.length)} de{" "}
                        {itemsProveedorFiltrados.length}
                      </div>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <DefaultButton
                          text="Anterior"
                          iconProps={{ iconName: "ChevronLeft" }}
                          onClick={() => setGridPage((prev) => Math.max(1, prev - 1))}
                          disabled={gridPage === 1}
                          styles={secondaryButtonStyles}
                        />
                        <div style={{ fontSize: 13, fontWeight: 600, color: BRAND.ink }}>
                          Página {gridPage} de {totalGridPages}
                        </div>
                        <DefaultButton
                          text="Siguiente"
                          iconProps={{ iconName: "ChevronRight" }}
                          onClick={() =>
                            setGridPage((prev) => Math.min(totalGridPages, prev + 1))
                          }
                          disabled={gridPage === totalGridPages}
                          styles={secondaryButtonStyles}
                        />
                      </Stack>
                    </Stack>
                  </Stack>
                )}
              </>
            )}
          </Stack>
        )}

        {/* Mensajes movidos al final del formulario */}
        {false && mensaje && (
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            styles={messageBarStyles}
          >
            {mensaje}
          </MessageBar>
        )}

        {/* ✅ Contenedor focuseable para enviar foco al error */}
        {false && error && (
          <div
            ref={errorRef}
            tabIndex={-1}
            role="alert"
            aria-live="assertive"
            aria-atomic="true"
            style={{ outline: "none" }}
          >
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
              styles={errorMessageBarStyles}
              messageBarIconProps={errorMessageBarIconProps}
            >
              {error}
            </MessageBar>
          </div>
        )}

        {/* ✅ Mostramos la validación en vivo sin mover el foco del campo editado */}
        {false && errorDocs && !error && (
          <div
            tabIndex={-1}
            role="alert"
            aria-live="assertive"
            aria-atomic="true"
            style={{ outline: "none" }}
          >
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
              styles={errorMessageBarStyles}
              messageBarIconProps={errorMessageBarIconProps}
            >
              {errorDocs}
            </MessageBar>
          </div>
        )}

        {false && modo === "Ingresar" && !docsObligIngresar && !errorDocs && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            isMultiline={true}
            styles={errorMessageBarStyles}
            messageBarIconProps={errorMessageBarIconProps}
          >
            DNI requiere fecha. Carnet de sanidad y los certificados (penales y policiales)
            requieren fecha y archivo. Si corresponde, la Licencia requiere fecha.
          </MessageBar>
        )}

        {/* Sección 2 - Datos personales */}
        <Stack
          tokens={{ childrenGap: 8 }}
          styles={sectionCardStyles}
        >
          <Label styles={sectionTitleStyles}>
            Datos personales
          </Label>

          <input type="hidden" name="ProveedorTitle" value={proveedorTitleOculto} />

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              {bloquearEmpresa ? (
                <TextField
                  label="Empresa"
                  value={proveedorTitleOculto || ""}
                  disabled
                  styles={roundedField}
                />
              ) : (
                <Dropdown
                  label="Empresa"
                  placeholder="Seleccionar empresa…"
                  options={proveedoresOptions}
                  selectedKey={proveedorId ?? undefined}
                  required
                  onChange={(_, opt) => {
                    touch();
                    const id = opt ? Number(opt.key) : null;
                    setProveedorId(id);
                    setProveedorTitleOculto(opt?.text ? String(opt.text) : "");
                  }}
                  styles={roundedDropdown}
                  disabled={isDarDeBaja}
                />
              )}
            </StackItem>
          </Stack>

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Nombre"
                value={form.Nombre}
                onChange={(_, v) => onChange("Nombre", v || "")}
                required={modo !== "Dar de baja"}
                styles={roundedField}
                disabled={isDarDeBaja}
              />
            </StackItem>
          </Stack>

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Apellido paterno"
                value={form.ApellidoPaterno}
                onChange={(_, v) => onChange("ApellidoPaterno", v || "")}
                styles={roundedField}
                disabled={isDarDeBaja}
              />
            </StackItem>
          </Stack>

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Apellido materno"
                value={form.ApellidoMaterno}
                onChange={(_, v) => onChange("ApellidoMaterno", v || "")}
                styles={roundedField}
                disabled={isDarDeBaja}
              />
            </StackItem>
          </Stack>

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <Dropdown
                label="Tipo de documento"
                options={opcionesTipoDocumento}
                selectedKey={form.TipoDocumento}
                onChange={(_, opt) => onChange("TipoDocumento", String(opt?.key))}
                styles={roundedDropdown}
                disabled={isDarDeBaja}
              />
            </StackItem>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Documento"
                value={form.Documento}
                onChange={(_, v) => onChange("Documento", v || "")}
                description={documentoAyuda}
                errorMessage={documentoErrorMessage}
                required
                styles={roundedField}
                disabled={isDarDeBaja}
              />
            </StackItem>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <span />
            </StackItem>
          </Stack>
        </Stack>

        {/* Datos laborales */}
        <Stack
          tokens={{ childrenGap: 8 }}
          styles={sectionCardStyles}
        >
          <Label styles={sectionTitleStyles}>
            Datos laborales
          </Label>

          <input type="hidden" name="ProveedorTitle" value={proveedorTitleOculto} />

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <Dropdown
                label="Puesto"
                options={opcionesPuesto}
                selectedKey={form.Puesto}
                onChange={(_, opt) => {
                  touch();
                  const nuevo = String(opt?.key || "");
                  const esConductor = nuevo.toLowerCase() === "conductor";
                  if (!esConductor) {
                    setLicCaducidad(null);
                    setLicFile(null);
                  }
                  setForm((prev) => ({
                    ...prev,
                    Puesto: nuevo,
                    Especificar: nuevo.toLowerCase() === "otro" ? prev.Especificar : "",
                    Licencia: esConductor ? prev.Licencia : "",
                    Categoria: esConductor ? prev.Categoria : undefined,
                  }));
                }}
                styles={roundedDropdown}
                disabled={isDarDeBaja || laboralBloqueado}
              />
            </StackItem>

            {showEspecificar && (
              <StackItem grow styles={{ root: { minWidth: 200 } }}>
                <TextField
                  label="Especificar (otro)"
                  value={form.Especificar}
                  onChange={(_, v) => onChange("Especificar", v || "")}
                  styles={roundedField}
                  disabled={isDarDeBaja || laboralBloqueado}
                />
              </StackItem>
            )}
          </Stack>

          {showLicenciaCat && (
            <Stack horizontal wrap tokens={stackTokens}>
              <StackItem grow styles={{ root: { minWidth: 200 } }}>
                <TextField
                  label="Licencia"
                  value={form.Licencia}
                  onChange={(_, v) => onChange("Licencia", v || "")}
                  styles={roundedField}
                  disabled={isDarDeBaja || laboralBloqueado}
                />
              </StackItem>
              <StackItem grow styles={{ root: { minWidth: 200 } }}>
                <Dropdown
                  label="Categoría"
                  options={opcionesCategoria}
                  selectedKey={form.Categoria}
                  onChange={(_, opt) => onChange("Categoria", String(opt?.key))}
                  styles={roundedDropdown}
                  disabled={isDarDeBaja || laboralBloqueado}
                />
              </StackItem>
            </Stack>
          )}
        </Stack>

        {/* Sección 3 - Documentación */}
        <Stack
          tokens={{ childrenGap: 12 }}
          styles={sectionCardStyles}
        >
          <Label styles={sectionTitleStyles}>
            Documentación
          </Label>

          {modo === "Ingresar" && (
            <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
              <DocCard
                title="DNI"
                dateLabel="Fecha de caducidad"
                dateValue={dniCaducidad}
                onDateChange={setDateAndTouch(setDniCaducidad)}
                minDate={fechaHoy}
                file={dniFile}
                onFileChange={setFileAndTouch(setDniFile)}
              />

              {showLicenciaCat && (
                <DocCard
                  title="Licencia"
                  dateLabel="Fecha de caducidad"
                  dateValue={licCaducidad}
                  onDateChange={setDateAndTouch(setLicCaducidad)}
                  minDate={fechaHoy}
                  file={licFile}
                  onFileChange={setFileAndTouch(setLicFile)}
                />
              )}

              <DocCard
                title="Carnet de sanidad"
                dateLabel="Fecha de emisión"
                dateValue={carnetEmision}
                onDateChange={setDateAndTouch(setCarnetEmision)}
                maxDate={fechaHoy}
                file={carnetFile}
                onFileChange={setFileAndTouch(setCarnetFile)}
              />
              <DocCard
                title="Antecedentes penales"
                dateLabel="Fecha de emisión"
                dateValue={penalesEmision}
                onDateChange={setDateAndTouch(setPenalesEmision)}
                maxDate={fechaHoy}
                file={penalesFile}
                onFileChange={setFileAndTouch(setPenalesFile)}
              />
              <DocCard
                title="Antecedentes policiales"
                dateLabel="Fecha de emisión"
                dateValue={policialesEmision}
                onDateChange={setDateAndTouch(setPolicialesEmision)}
                maxDate={fechaHoy}
                file={policialesFile}
                onFileChange={setFileAndTouch(setPolicialesFile)}
              />
            </Stack>
          )}

          {modo === "Modificar" && (
            <Stack tokens={{ childrenGap: 8 }}>
              {!form.Documento?.trim() ? (
                <div style={infoBannerStyles.root}>
                  <Icon
                    iconName="Info"
                    styles={{ root: infoBannerStyles.icon }}
                  />
                  <div style={infoBannerStyles.text}>
                    Seleccioná un registro en la grilla superior para ver su documentación.
                  </div>
                </div>
              ) : (
                <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                  {docRows.map((r) => (
                    <DocCard
                      key={r.key}
                      title={r.label}
                      dateLabel={r.tipo === "cad" ? "Fecha de caducidad" : "Fecha de emisión"}
                      dateValue={r.fecha}
                      onDateChange={(d) => {
                        touch();
                        setDocRows((prev) =>
                          prev.map((row) =>
                            row.label === r.label ? { ...row, fecha: d } : row
                          )
                        );
                      }}
                      minDate={r.tipo === "cad" ? fechaHoy : undefined}
                      maxDate={r.tipo === "emi" ? fechaHoy : undefined}
                      file={r.file || null}
                      onFileChange={(file) => {
                        touch();
                        setDocRows((prev) =>
                          prev.map((row) =>
                            row.label === r.label ? { ...row, file } : row
                          )
                        );
                      }}
                      attachments={r.attachments}
                    />
                  ))}
                </Stack>
              )}
            </Stack>
          )}
        </Stack>

        {/* Sección 4 - Notificaciones */}
        <Stack
          tokens={{ childrenGap: 8 }}
          styles={sectionCardStyles}
        >
          <Label styles={sectionTitleStyles}>
            Notificaciones
          </Label>
          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Correos de notificación"
                placeholder="correo1@dominio.com; correo2@dominio.com"
                value={form.CorreosNotificacion}
                onChange={(_, v) => onChange("CorreosNotificacion", v || "")}
                multiline
                autoAdjustHeight
                styles={roundedField}
                disabled={isDarDeBaja}
              />
            </StackItem>
          </Stack>
        </Stack>

        {/* Modal de confirmación para Dar de baja */}
        {/* Mensajes entre notificaciones y acciones */}
        {mensaje && (
          <div
            role="status"
            aria-live="polite"
            aria-atomic="true"
          >
            <div style={successBannerStyles.root}>
              <Icon
                iconName="Completed"
                styles={{ root: successBannerStyles.icon }}
              />
              <div style={successBannerStyles.text}>{mensaje}</div>
            </div>
          </div>
        )}

        {/* Contenedor focuseable para enviar foco al error */}
        {error && (
          <div
            ref={errorRef}
            tabIndex={-1}
            role="alert"
            aria-live="assertive"
            aria-atomic="true"
            style={{ outline: "none" }}
          >
            <div style={dangerBannerStyles.root}>
              <Icon
                iconName="StatusErrorFull"
                styles={{ root: dangerBannerStyles.icon }}
              />
              <div style={dangerBannerStyles.text}>{error}</div>
            </div>
          </div>
        )}

        {/* Mostramos la validacion en vivo sin mover el foco del campo editado */}
        {errorDocs && !error && (
          <div
            tabIndex={-1}
            role="alert"
            aria-live="assertive"
            aria-atomic="true"
            style={{ outline: "none" }}
          >
            <div style={dangerBannerStyles.root}>
              <Icon
                iconName="StatusErrorFull"
                styles={{ root: dangerBannerStyles.icon }}
              />
              <div style={dangerBannerStyles.text}>{errorDocs}</div>
            </div>
          </div>
        )}

        {modo === "Ingresar" && requiredFieldsPendientesIngresar.length > 0 && (
          <div
            role="alert"
            aria-live="assertive"
            aria-atomic="true"
            style={dangerBannerStyles.root}
          >
            <Icon
              iconName="StatusErrorFull"
              styles={{ root: dangerBannerStyles.icon }}
            />
            <div style={dangerBannerStyles.text}>
              <div style={requiredFieldsListStyles.title}>
                Campos obligatorios pendientes:
              </div>
              <ul style={requiredFieldsListStyles.list}>
                {requiredFieldsPendientesIngresar.map((campo) => (
                  <li key={campo} style={requiredFieldsListStyles.item}>
                    {campo}
                  </li>
                ))}
              </ul>
            </div>
          </div>
        )}

        {/* Modal de confirmacion para Dar de baja */}
        <Dialog
          hidden={!showConfirmBaja}
          onDismiss={() => {
            if (!guardando) {
              setShowConfirmBaja(false);
              setMotivoBaja("");
            }
          }}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Confirmar baja",
            subText: borrar
              ? "Se eliminará el registro y toda la documentación."
              : "El registro se marcará como inactivo.",
          }}
        >
          <TextField
            label="Motivo"
            multiline
            required
            value={motivoBaja}
            onChange={(_, v) => {
              touch();
              setMotivoBaja(v || "");
            }}
            styles={roundedField}
          />

          <DialogFooter>
            <PrimaryButton
              text="Confirmar"
              onClick={onConfirmarBaja}
              disabled={!motivoBaja.trim() || guardando}
              styles={primaryButtonStyles}
            />
            <DefaultButton
              text="Cancelar"
              onClick={() => {
                if (!guardando) {
                  setShowConfirmBaja(false);
                  setMotivoBaja("");
                }
              }}
              disabled={guardando}
              styles={secondaryButtonStyles}
            />
          </DialogFooter>
        </Dialog>

        {/* Sección 5 - Acciones */}
        <Stack tokens={{ childrenGap: 12 }} styles={sectionCardStyles}>
          <Label styles={sectionTitleStyles}>Acciones</Label>
          <Stack horizontal wrap tokens={stackTokens} verticalAlign="center">
            <PrimaryButton
              text="Guardar"
              iconProps={{ iconName: "Save" }}
              onClick={onGuardar}
              disabled={!puedeGuardar}
              styles={primaryButtonStyles}
            />
            <DefaultButton
              text="Cancelar"
              iconProps={{ iconName: "Clear" }}
              onClick={onCancelar}
              disabled={guardando}
              styles={secondaryButtonStyles}
            />
            {guardando && (
              <StackItem grow styles={{ root: { minWidth: 220 } }}>
                <ProgressIndicator label="Guardando..." />
              </StackItem>
            )}
          </Stack>
        </Stack>
      </Stack>
    </ThemeProvider>
  );
};

export default RegistroPersonal;
