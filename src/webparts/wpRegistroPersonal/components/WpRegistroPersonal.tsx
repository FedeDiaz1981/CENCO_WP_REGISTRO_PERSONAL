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
  ChoiceGroup,
  IChoiceGroupOption,
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

const LST_PERSONAS = "Personal";
const LST_DOCS = "Documentacion";

const stackTokens = { childrenGap: 12 };
const esc = (s: string) => s.replace(/'/g, "''");
const dateToISO = (d?: Date | null) => (d ? d.toISOString() : null);

type DocFields = { Caducidad?: string | null; Emision?: string | null };

// ---- Estilos / Tema ----
const theme = createTheme({
  palette: {
    themePrimary: "#2563eb",
    themeLighterAlt: "#f6f9ff",
    themeLighter: "#d8e6fe",
    themeLight: "#b6d0fd",
    themeTertiary: "#6ea4fb",
    themeSecondary: "#377df7",
    themeDarkAlt: "#2159d3",
    themeDark: "#1c49b0",
    themeDarker: "#14357f",
    neutralLighterAlt: "#faf9f8",
    neutralLighter: "#f3f2f1",
    neutralLight: "#edebe9",
    neutralQuaternaryAlt: "#e1dfdd",
    neutralQuaternary: "#d0d0d0",
    neutralTertiaryAlt: "#c8c6c4",
    neutralTertiary: "#a19f9d",
    neutralSecondary: "#605e5c",
    neutralPrimaryAlt: "#3b3a39",
    neutralPrimary: "#323130",
    neutralDark: "#201f1e",
    black: "#000000",
    white: "#ffffff",
  },
  effects: {
    roundedCorner2: "12px",
    elevation8: "0 6px 18px rgba(0,0,0,.08)" as any,
  },
});

// ---- Mini componente: Tarjeta documento (para INGRESAR) ----
interface DocCardProps {
  title: string;
  dateLabel: string;
  dateValue: Date | null;
  onDateChange: (date: Date | null) => void;
  file: File | null;
  onFileChange: (file: File | null) => void;
}


const DocCard: React.FC<DocCardProps> = ({
  title,
  dateLabel,
  dateValue,
  onDateChange,
  file,
  onFileChange,
}) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  return (
    <Stack
      horizontal
      wrap
      verticalAlign="end"
      tokens={{ childrenGap: 12 }}
      styles={{
        root: {
          border: `1px solid ${theme.palette.neutralLight}`,
          borderRadius: 12,
          padding: 12,
          boxShadow: theme.effects.elevation8,
          background: theme.palette.white,
        },
      }}
    >
      <StackItem styles={{ root: { minWidth: 160 } }}>
        <Label styles={{ root: { fontWeight: 600 } }}>{title}</Label>
      </StackItem>
      <StackItem grow styles={{ root: { minWidth: 220, maxWidth: 320 } }}>
        <DatePicker
          label={dateLabel}
          value={dateValue || undefined}
          onSelectDate={(d) => onDateChange(d ?? null)}
          firstDayOfWeek={DayOfWeek.Monday}
          placeholder="Seleccionar fecha"
          ariaLabel={dateLabel}
        />
      </StackItem>
      <StackItem grow styles={{ root: { minWidth: 220, maxWidth: 340 } }}>
        <Label>Archivo adjunto</Label>
        <input
          ref={fileInputRef}
          type="file"
          style={{ display: "none" }}
          onChange={(e) =>
            onFileChange(
              e.target.files && e.target.files.length ? e.target.files[0] : null
            )
          }
        />
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <DefaultButton
            iconProps={{ iconName: "Upload" }}
            text={file ? "Cambiar archivo" : "Adjuntar"}
            onClick={() => fileInputRef.current?.click()}
          />
          {file && (
            <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
              <Icon iconName="Page" />
              <span style={{ wordBreak: "break-all" }}>{file.name}</span>
              <DefaultButton
                text="Quitar"
                onClick={() => onFileChange(null)}
                styles={{ root: { marginLeft: 6 } }}
              />
            </Stack>
          )}
        </Stack>
      </StackItem>
    </Stack>
  );
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

const cutoffSinceMonths = (months: number) => addMonthsSafe(today0(), -months);

// ===========================================================
// REGISTRO DE PERSONAL
// ===========================================================
const RegistroPersonal: React.FC<IRegistroPersonalProps> = ({ sp, siteUrl, filtrarPorProveedor }) => {
  const [modo, setModo] = React.useState<Modo>("Ingresar");
  const [proveedorTitleOculto, setProveedorTitleOculto] = React.useState("");
  const [proveedorId, setProveedorId] = React.useState<number | null>(null);

  // -------- Sección 2: Personal --------
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

  // ===== visibilidad dinámica por Puesto =====
  const puestoNorm = (form.Puesto || "").toLowerCase().trim();
  const showEspecificar = React.useMemo(() => puestoNorm === "otro", [puestoNorm]);
  const showLicenciaCat = React.useMemo(() => puestoNorm === "conductor", [puestoNorm]);

  // -------- Sección 3: Documentación (para Ingresar) --------
  // IMPORTANTE: estos estados van ANTES de errorDocs
  const [dniCaducidad, setDniCaducidad] = React.useState<Date | null>(null);
  const [dniFile, setDniFile] = React.useState<File | null>(null);
  const [licCaducidad, setLicCaducidad] = React.useState<Date | null>(null);
  const [licFile, setLicFile] = React.useState<File | null>(null);
  const [carnetEmision, setCarnetEmision] = React.useState<Date | null>(null);
  const [carnetFile, setCarnetFile] = React.useState<File | null>(null);
  const [penalesEmision, setPenalesEmision] = React.useState<Date | null>(null);
  const [penalesFile, setPenalesFile] = React.useState<File | null>(null);
  const [policialesEmision, setPolicialesEmision] = React.useState<Date | null>(null);
  const [policialesFile, setPolicialesFile] = React.useState<File | null>(null);

  // ---- Estado de grilla de documentación (Modificar) ----
  const [docRows, setDocRows] = React.useState<DocRow[]>(makeDefaultDocRows());

  const [guardando, setGuardando] = React.useState(false);
  const [mensaje, setMensaje] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);

  // ---- Validación de antigüedad ----
  const errorDocs = React.useMemo(() => {
    const fmt = (d: Date) => d.toLocaleDateString();

    if (modo === "Ingresar") {
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
  // util local sin .find()
  const getFechaByLabel = (lbl: string): Date | null => {
    for (let i = 0; i < docRows.length; i++) {
      if (docRows[i].label === lbl) return docRows[i].fecha || null;
    }
    return null;
  };

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
  }, [modo, carnetEmision, penalesEmision, policialesEmision, docRows]);

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
    { key: "ap", name: "Ap. paterno", fieldName: "Apellido_x0020_paterno", minWidth: 120, isResizable: true },
    { key: "am", name: "Ap. materno", fieldName: "Apellido_x0020_materno", minWidth: 120, isResizable: true },
    { key: "pto", name: "Puesto", fieldName: "puesto", minWidth: 120, isResizable: true },
    { key: "cat", name: "Categoría", fieldName: "Categoria", minWidth: 90, isResizable: true },
  ];

  const [itemsProveedor, setItemsProveedor] = React.useState<PersonaItem[]>([]);
  const [cargandoGrid, setCargandoGrid] = React.useState(false);

  const [queryGrid, setQueryGrid] = React.useState("");
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
          "correosnotificacion"
        )
        .orderBy("Id", false)
        .top(5000);

      // si el toggle está activo, filtra por el Proveedor del usuario
      if (filtrarPorProveedor) {
        if (!proveedorId) {
          if (!cancelado) setItemsProveedor([]);
          return;
        }
        query = query.filter(`ProveedorId eq ${proveedorId}`);
      }

      const items = await query();
      if (!cancelado) setItemsProveedor(items as PersonaItem[]);
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
}, [modo, proveedorId, sp, filtrarPorProveedor]);

  // Meta lookup Proveedor
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



  // Carga meta campo + proveedor del usuario
  React.useEffect(() => {
    let cancelado = false;
    const cargar = async () => {
      try {
        const f = await sp.web.lists
          .getByTitle(LST_PERSONAS)
          .fields.getByInternalNameOrTitle("Proveedor")
          .select("InternalName", "TypeAsString", "AllowMultipleValues")();
        if (!cancelado)
          setProvFieldMeta({
            InternalName: f.InternalName,
            TypeAsString: f.TypeAsString,
          });

        const me = await sp.web.currentUser();
        let items = await sp.web.lists
          .getByTitle("Proveedores")
          .items.select("Id", "Title", "Usuarios/Id")
          .expand("Usuarios")
          .filter(`Usuarios/Id eq ${me.Id}`)
          .top(1)();
        if (!items?.length) {
          items = await sp.web.lists
            .getByTitle("Proveedores")
            .items.select("Id", "Title", "UsuariosId")
            .filter(`UsuariosId eq ${me.Id}`)
            .top(1)();
        }
        if (!cancelado && items?.[0]) {
          setProveedorTitleOculto(items[0].Title);
          setProveedorId(items[0].Id);
        }
      } catch {}
    };
    cargar().catch(() => {});
    return () => {
      cancelado = true;
    };
  }, [sp]);

  const onChange = (field: keyof PersonaForm, value?: string) =>
    setForm((prev) => ({ ...prev, [field]: value ?? "" }));

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

  const limpiar = () => {
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
    setMensaje(null);
    setError(null);
    setDocRows(makeDefaultDocRows());
  };

  // ----------------- Utilidades Personas -----------------
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
      correosnotificacion: form.CorreosNotificacion,
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
        correosnotificacion: form.CorreosNotificacion,
        ...buildProveedorPayload(),
      });

  const eliminarEnPersonas = async (id: number) =>
    sp.web.lists.getByTitle(LST_PERSONAS).items.getById(id).delete();

  // ----------------- Utilidades Documentación -----------------
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
      .filter(
        `Title eq '${esc(form.Documento)}' and Documento eq '${esc(label)}'`
      )
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
          } catch {}
        }
      }
    } catch {}
    await item.attachmentFiles.add(file.name, file);
  };

  // Inserta o actualiza sin borrar adjuntos existentes
  const upsertDocRow = async (label: string, fields: DocFields, file?: File | null) => {
    if (!form.Documento?.trim())
      throw new Error("Documento (Title) es obligatorio para Documentación.");
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

  // ----------------- Guardar (según modo) -----------------
  const onGuardar = async () => {
    setMensaje(null);
    setError(null);

    if (!form.Documento?.trim()) {
      setError("Documento es obligatorio.");
      return;
    }
    if (modo !== "Dar de baja" && !form.Nombre?.trim()) {
      setError("Nombre es obligatorio.");
      return;
    }
    if (errorDocs) {
      setError(errorDocs);
      return;
    }

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
        if (persona) await eliminarEnPersonas(persona.Id);
        await deleteAllDocsByTitle(form.Documento);
        setMensaje("Registro dado de baja en Personas y toda la Documentación.");
      }

      setDniFile(null);
    } catch (e: any) {
      setError(e.message ?? "Error al guardar.");
    } finally {
      setGuardando(false);
    }
  };

  function toDate(iso?: string | null): Date | null {
    if (!iso) return null;
    const d = new Date(iso);
    return isNaN(d.getTime()) ? null : d;
  }

  // ====== Cargar documentación (fechas + grilla de docs con adjuntos) ======
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
      { key: "Carnet de sanidad", tipo: "emi" as const, fechaRaw: map.get("Carnet de sanidad")?.Emision ?? null },
      { key: "Antecedentes penales", tipo: "emi" as const, fechaRaw: map.get("Antecedentes penales")?.Emision ?? null },
      { key: "Antecedentes policiales", tipo: "emi" as const, fechaRaw: map.get("Antecedentes policiales")?.Emision ?? null },
    ];

    const withAtts: DocRow[] = [];
    for (let i = 0; i < defs.length; i++) {
      const d = defs[i];
      const r = map.get(d.key);
      let attachments: Attach[] = [];
      if (r?.Id) {
        try {
          const atts = await sp.web.lists.getByTitle(LST_DOCS).items.getById(r.Id).attachmentFiles();
          attachments = (atts || []).map((a: any) => ({
            name: a.FileName,
            href: toAbs(siteUrl, a.ServerRelativeUrl),
          }));
        } catch {}
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

  // Refrescar sólo una fila tras subir
  async function refreshDocRow(label: string, docTitle: string) {
    const it = await getDocItemByLabel(docTitle, label);
    let attachments: Attach[] = [];
    if (it?.Id) {
      const atts = await sp.web.lists.getByTitle(LST_DOCS).items.getById(it.Id).attachmentFiles();
      attachments = (atts || []).map((a: any) => ({
        name: a.FileName,
        href: toAbs(siteUrl, a.ServerRelativeUrl),
      }));
    }
    setDocRows((prev) =>
      prev.map((r) =>
        r.label === label
          ? { ...r, attachments, editing: false, file: null, justUpdated: true }
          : r
      )
    );
    setTimeout(
      () =>
        setDocRows((prev) =>
          prev.map((r) => (r.label === label ? { ...r, justUpdated: false } : r))
        ),
      3000
    );
  }

  // Subida inmediata por etiqueta
  async function uploadForRowByLabel(label: string) {
    if (!form.Documento?.trim()) {
      setError("Documento (Title) es obligatorio.");
      return;
    }
    let cur: DocRow | undefined = undefined;
    for (let i = 0; i < docRows.length; i++) {
      if (docRows[i].label === label) {
        cur = docRows[i];
        break;
      }
    }
    if (!cur) {
      setError("No se encontró la fila de documentación.");
      return;
    }
    if (!cur.file) {
      setError("Seleccioná un archivo para subir.");
      return;
    }

    const fields =
      cur.tipo === "cad" ? { Caducidad: dateToISO(cur.fecha) } : { Emision: dateToISO(cur.fecha) };

    try {
      await upsertDocRow(cur.label, fields, cur.file);
      await refreshDocRow(cur.label, form.Documento);
    } catch (e) {
      console.warn("uploadForRow error:", e);
    }
  }

  // Cargar desde la grilla de personas
  async function loadFromGridItem(it: PersonaItem) {
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
      CorreosNotificacion: it.correosnotificacion ?? "",
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

  const modoOptions: IChoiceGroupOption[] = [
    { key: "Ingresar", text: "Ingresar", iconProps: { iconName: "Add" } },
    { key: "Modificar", text: "Modificar", iconProps: { iconName: "Edit" } },
    { key: "Dar de baja", text: "Dar de baja", iconProps: { iconName: "Delete" } },
  ];

  return (
    <ThemeProvider theme={theme}>
      <Stack tokens={stackTokens} styles={{ root: { maxWidth: 1024, margin: "0 auto", padding: 12 } }}>
        {/* Barra de modo */}
        <Stack
          horizontal
          wrap
          verticalAlign="center"
          tokens={{ childrenGap: 16 }}
          styles={{
            root: {
              background: theme.palette.white,
              borderRadius: 12,
              padding: 12,
              boxShadow: theme.effects.elevation8 as any,
            },
          }}
        >
          <Icon iconName="Contact" styles={{ root: { fontSize: 22, color: theme.palette.themePrimary } }} />
          <Label styles={{ root: { fontSize: 18, fontWeight: 600 } }}>Registro de Personal</Label>
          <StackItem grow />
          <ChoiceGroup
            selectedKey={modo}
            options={modoOptions}
            onChange={(_, opt) => {
              const next = (opt?.key as Modo) ?? "Ingresar";
              setModo(next);
              if (next === "Ingresar") {
                limpiar();
                try {
                  selectionRef.current?.setAllSelected(false);
                } catch {}
              }
            }}
          />
        </Stack>

        {/* Grilla de personas del proveedor */}
        {(modo === "Modificar" || modo === "Dar de baja") && (
          <Stack
            tokens={{ childrenGap: 8 }}
            styles={{
              root: {
                background: theme.palette.white,
                borderRadius: 12,
                padding: 12,
                boxShadow: theme.effects.elevation8 as any,
              },
            }}
          >
            <Label styles={{ root: { fontWeight: 600 } }}>
              Registros del proveedor: {proveedorTitleOculto || "(sin proveedor)"}
            </Label>

            {cargandoGrid ? (
              <Spinner label="Cargando registros..." size={SpinnerSize.small} />
            ) : (
              <>
                <SearchBox
                  placeholder="Filtrar por cualquier campo…"
                  value={queryGrid}
                  onChange={(_, v) => setQueryGrid(v || "")}
                  onClear={() => setQueryGrid("")}
                  styles={{ root: { maxWidth: 340, marginBottom: 8 } }}
                />

                {itemsProveedorFiltrados.length === 0 ? (
                  <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
                    {queryGrid ? "Sin resultados para la búsqueda." : "No hay registros."}
                  </MessageBar>
                ) : (
                  <div style={{ width: "100%", overflowX: "auto" }}>
                    <DetailsList
                      items={itemsProveedorFiltrados}
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
                )}
              </>
            )}
          </Stack>
        )}

        {/* Mensajes */}
        {mensaje && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            {mensaje}
          </MessageBar>
        )}
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
            {error}
          </MessageBar>
        )}
        {errorDocs && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
            {errorDocs}
          </MessageBar>
        )}

        {/* Sección 2 - Personal */}
        <Stack
          tokens={{ childrenGap: 8 }}
          styles={{
            root: {
              background: theme.palette.white,
              borderRadius: 12,
              padding: 16,
              boxShadow: theme.effects.elevation8 as any,
            },
          }}
        >
          <Label styles={{ root: { fontWeight: 600, fontSize: 16 } }}>Datos personales</Label>
          <input type="hidden" name="ProveedorTitle" value={proveedorTitleOculto} />

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Nombre"
                value={form.Nombre}
                onChange={(_, v) => onChange("Nombre", v || "")}
                required={modo !== "Dar de baja"}
              />
            </StackItem>
          </Stack>

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Apellido paterno"
                value={form.ApellidoPaterno}
                onChange={(_, v) => onChange("ApellidoPaterno", v || "")}
              />
            </StackItem>
          </Stack>

          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Apellido materno"
                value={form.ApellidoMaterno}
                onChange={(_, v) => onChange("ApellidoMaterno", v || "")}
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
              />
            </StackItem>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Documento (Title)"
                value={form.Documento}
                onChange={(_, v) => onChange("Documento", v || "")}
                required
              />
            </StackItem>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <span />
            </StackItem>
          </Stack>

          {/* Puesto + campos condicionales */}
          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <Dropdown
                label="Puesto"
                options={opcionesPuesto}
                selectedKey={form.Puesto}
                onChange={(_, opt) => {
                  const nuevo = String(opt?.key || "");
                  setForm((prev) => ({
                    ...prev,
                    Puesto: nuevo,
                    Especificar: nuevo.toLowerCase() === "otro" ? prev.Especificar : "",
                    Licencia: nuevo.toLowerCase() === "conductor" ? prev.Licencia : "",
                    Categoria: nuevo.toLowerCase() === "conductor" ? prev.Categoria : undefined,
                  }));
                }}
              />
            </StackItem>

            {showEspecificar && (
              <StackItem grow styles={{ root: { minWidth: 200 } }}>
                <TextField
                  label="Especificar (otro)"
                  value={form.Especificar}
                  onChange={(_, v) => onChange("Especificar", v || "")}
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
                />
              </StackItem>
              <StackItem grow styles={{ root: { minWidth: 200 } }}>
                <Dropdown
                  label="Categoría"
                  options={opcionesCategoria}
                  selectedKey={form.Categoria}
                  onChange={(_, opt) => onChange("Categoria", String(opt?.key))}
                />
              </StackItem>
            </Stack>
          )}
        </Stack>

        {/* Sección 2.5 - Notificaciones */}
        <Stack
          tokens={{ childrenGap: 8 }}
          styles={{
            root: {
              background: theme.palette.white,
              borderRadius: 12,
              padding: 16,
              boxShadow: theme.effects.elevation8 as any,
            },
          }}
        >
          <Label styles={{ root: { fontWeight: 600, fontSize: 16 } }}>Notificaciones</Label>
          <Stack horizontal wrap tokens={stackTokens}>
            <StackItem grow styles={{ root: { minWidth: 200 } }}>
              <TextField
                label="Correos de notificación"
                placeholder="correo1@dominio.com; correo2@dominio.com"
                value={form.CorreosNotificacion}
                onChange={(_, v) => onChange("CorreosNotificacion", v || "")}
                multiline
                autoAdjustHeight
              />
            </StackItem>
          </Stack>
        </Stack>

        {/* Sección 3 - Documentación */}
        <Stack tokens={{ childrenGap: 12 }}>
          <Label styles={{ root: { fontWeight: 600, fontSize: 16 } }}>Documentación</Label>

          {modo === "Ingresar" && (
            <>
              <DocCard
                title="DNI"
                dateLabel="Caducidad"
                dateValue={dniCaducidad}
                onDateChange={setDniCaducidad}
                file={dniFile}
                onFileChange={setDniFile}
              />

              {showLicenciaCat && (
                <DocCard
                  title="Licencia"
                  dateLabel="Caducidad"
                  dateValue={licCaducidad}
                  onDateChange={setLicCaducidad}
                  file={licFile}
                  onFileChange={setLicFile}
                />
              )}

              <DocCard
                title="Carnet de sanidad"
                dateLabel="Fecha de emisión"
                dateValue={carnetEmision}
                onDateChange={setCarnetEmision}
                file={carnetFile}
                onFileChange={setCarnetFile}
              />
              <DocCard
                title="Antecedentes penales"
                dateLabel="Fecha de emisión"
                dateValue={penalesEmision}
                onDateChange={setPenalesEmision}
                file={penalesFile}
                onFileChange={setPenalesFile}
              />
              <DocCard
                title="Antecedentes policiales"
                dateLabel="Fecha de emisión"
                dateValue={policialesEmision}
                onDateChange={setPolicialesEmision}
                file={policialesFile}
                onFileChange={setPolicialesFile}
              />
            </>
          )}

          {modo === "Modificar" && (
            <Stack tokens={{ childrenGap: 8 }}>
              {!form.Documento?.trim() && (
                <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
                  Seleccioná un registro en la grilla superior para ver su documentación.
                </MessageBar>
              )}
              <DetailsList
                items={docRows}
                selectionMode={SelectionMode.none}
                columns={[
                  {
                    key: "colLabel",
                    name: "Documento",
                    minWidth: 180,
                    onRender: (r: DocRow) => (
                      <span style={{ color: r.justUpdated ? "green" : undefined }}>{r.label}</span>
                    ),
                  },
                  {
                    key: "colFecha",
                    name: "Fecha",
                    minWidth: 160,
                    onRender: (r: DocRow) => (
                      <span style={{ color: r.justUpdated ? "green" : undefined }}>
                        {r.fecha ? r.fecha.toLocaleDateString() : "-"}{" "}
                        <i>({r.tipo === "cad" ? "Caducidad" : "Emisión"})</i>
                      </span>
                    ),
                  },
                  {
                    key: "colAdj",
                    name: "Adjuntos",
                    minWidth: 240,
                    onRender: (r: DocRow) =>
                      r.attachments?.length ? (
                        <Stack tokens={{ childrenGap: 4 }}>
                          {r.attachments.map((a) => (
                            <a key={a.href} href={a.href} target="_blank" rel="noopener noreferrer">
                              {a.name}
                            </a>
                          ))}
                        </Stack>
                      ) : (
                        <span style={{ opacity: 0.6 }}>Sin adjuntos</span>
                      ),
                  },
                  {
                    key: "colAccion",
                    name: "Acción",
                    minWidth: 340,
                    onRender: (r: DocRow) =>
                      r.editing ? (
                        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                          <input
                            type="file"
                            onChange={(e) => {
                              const file = e.target.files && e.target.files[0] ? e.target.files[0] : null;
                              setDocRows((prev) => {
                                const next = prev.slice(0);
                                for (let i = 0; i < next.length; i++) {
                                  if (next[i].label === r.label) {
                                    next[i] = { ...next[i], file };
                                    break;
                                  }
                                }
                                return next;
                              });
                            }}
                          />
                          <PrimaryButton
                            text="Subir"
                            disabled={!r.file}
                            onClick={() =>
                              uploadForRowByLabel(r.label).catch((err) => {
                                console.warn(err);
                                setError("No se pudo subir el adjunto.");
                              })
                            }
                          />
                          <DefaultButton
                            text="Confirmar"
                            onClick={() =>
                              setDocRows((prev) => {
                                const next = prev.slice(0);
                                for (let i = 0; i < next.length; i++) {
                                  if (next[i].label === r.label) {
                                    next[i] = { ...next[i], editing: false, justUpdated: true };
                                    break;
                                  }
                                }
                                setTimeout(() => {
                                  setDocRows((p2) => {
                                    const n2 = p2.slice(0);
                                    for (let j = 0; j < n2.length; j++)
                                      if (n2[j].label === r.label)
                                        n2[j] = { ...n2[j], justUpdated: false };
                                    return n2;
                                  });
                                }, 3000);
                                return next;
                              })
                            }
                          />
                          <DefaultButton
                            text="Cancelar"
                            onClick={() =>
                              setDocRows((prev) => {
                                const next = prev.slice(0);
                                for (let i = 0; i < next.length; i++) {
                                  if (next[i].label === r.label) {
                                    next[i] = { ...next[i], editing: false, file: null };
                                    break;
                                  }
                                }
                                return next;
                              })
                            }
                          />
                        </Stack>
                      ) : (
                        <DefaultButton
                          text="Editar"
                          onClick={() =>
                            setDocRows((prev) => {
                              const next = prev.slice(0);
                              for (let i = 0; i < next.length; i++) {
                                if (next[i].label === r.label) {
                                  next[i] = { ...next[i], editing: true };
                                  break;
                                }
                              }
                              return next;
                            })
                          }
                        />
                      ),
                  },
                ]}
              />
            </Stack>
          )}
        </Stack>

        {/* Sección 4 - Acciones */}
        <Stack horizontal wrap tokens={stackTokens} verticalAlign="center">
          <PrimaryButton text="Guardar" onClick={onGuardar} disabled={guardando || !!errorDocs || !!error} />
          <DefaultButton text="Cancelar" onClick={onCancelar} disabled={guardando} />
          {guardando && (
            <StackItem grow>
              <ProgressIndicator label="Guardando..." />
            </StackItem>
          )}
        </Stack>
      </Stack>
    </ThemeProvider>
  );
};

export default RegistroPersonal;
