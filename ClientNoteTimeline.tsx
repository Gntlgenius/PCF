import * as React from "react";

/* ─────────────────────────────────────────────
   Types
───────────────────────────────────────────── */
export interface ClientNote {
  id: string;
  acc_name: string;
  acc_ipfclientstatus: string;
  acc_ipfclientstatusLabel: string;
  ownerName: string;
  ownerInitials: string;
  ownerColor: string;
  position: string;
  description: string;
  acc_relatedto: string;
  acc_relatedtoLabel: string;
  createdon: string;
}

interface Props {
  clientId: string;
  webAPI: ComponentFramework.WebApi;
  navigation: ComponentFramework.Navigation;
  utils: ComponentFramework.Utility;
  containerWidth: number;
}

/* ─────────────────────────────────────────────
   Only one hardcoded colour — the fallback
───────────────────────────────────────────── */
const DEFAULT_BADGE_COLOR = "#95a5a6";

/* ─────────────────────────────────────────────
   Avatar colours keyed on name hash
───────────────────────────────────────────── */
const AVATAR_COLOURS = [
  "#3498db", "#2ecc71", "#e74c3c", "#9b59b6",
  "#f39c12", "#1abc9c", "#e67e22", "#34495e",
];

function avatarColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash += name.charCodeAt(i);
  return AVATAR_COLOURS[hash % AVATAR_COLOURS.length];
}

function initials(name: string): string {
  return name
    .split(" ")
    .slice(0, 2)
    .map((n) => n[0]?.toUpperCase() ?? "")
    .join("");
}

/* ─────────────────────────────────────────────
   Date formatter  dd/MM/yyyy
───────────────────────────────────────────── */
function fmtDate(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  return `${String(d.getDate()).padStart(2, "0")}/${String(
    d.getMonth() + 1
  ).padStart(2, "0")}/${d.getFullYear()}`;
}

/* ─────────────────────────────────────────────
   Fetch option set metadata via WebAPI
   Returns a map of { optionValue -> colorHex }
───────────────────────────────────────────── */
async function fetchOptionSetColors(
  webAPI: ComponentFramework.WebApi
): Promise<Record<string, string>> {
  try {
    // Query the OptionSetMetadata for acc_ipfclientstatus on acc_clientnote
    const result = await (webAPI as any).retrieveMultipleRecords(
      undefined,
      `/EntityDefinitions(LogicalName='acc_clientnote')/Attributes(LogicalName='acc_ipfclientstatus')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)`
    );

    const options: any[] =
      result?.OptionSet?.Options ??
      result?.entities?.[0]?.OptionSet?.Options ??
      [];

    const map: Record<string, string> = {};
    for (const opt of options) {
      const value = String(opt.Value ?? "");
      // Color comes back as an integer ARGB (e.g. 16711680 = #FF0000)
      // or as a hex string depending on the API version
      const raw = opt.Color ?? opt.color ?? null;
      if (raw !== null && raw !== undefined) {
        map[value] = typeof raw === "number"
          ? "#" + ((raw & 0xFFFFFF) >>> 0).toString(16).padStart(6, "0")
          : String(raw).startsWith("#") ? raw : `#${raw}`;
      }
    }
    return map;
  } catch {
    return {};
  }
}

/* ─────────────────────────────────────────────
   Main component
───────────────────────────────────────────── */
export const ClientNoteTimelineComponent: React.FC<Props> = ({
  clientId,
  webAPI,
  navigation,
  utils,
  containerWidth,
}) => {
  const [notes, setNotes] = React.useState<ClientNote[]>([]);
  const [filtered, setFiltered] = React.useState<ClientNote[]>([]);
  const [search, setSearch] = React.useState("");
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  // Map of optionset value → hex color from Dataverse metadata
  const [statusColorMap, setStatusColorMap] = React.useState<Record<string, string>>({});

  /* ── Fetch option set colours once on mount ── */
  React.useEffect(() => {
    fetchOptionSetColors(webAPI).then(setStatusColorMap);
  }, [webAPI]);

  /* ── Resolve badge colour: metadata first, fallback last ── */
  const getStatusColor = React.useCallback(
    (value: string): string =>
      statusColorMap[value] ?? DEFAULT_BADGE_COLOR,
    [statusColorMap]
  );

  /* ── Fetch notes ─────────────────────────── */
  const fetchNotes = React.useCallback(async () => {
    if (!clientId) return;
    setLoading(true);
    setError(null);
    try {
      const select = [
        "acc_clientnoteid",
        "acc_name",
        "acc_ipfclientstatus",
        "description",
        "createdon",
        "_ownerid_value",
        "acc_relatedto",
      ].join(",");

      const filter = `acc_clientid eq '${clientId}'`;
      const orderby = "createdon desc";

      const result = await webAPI.retrieveMultipleRecords(
        "acc_clientnote",
        `?$select=${select}&$filter=${filter}&$orderby=${orderby}&$expand=ownerid($select=fullname,jobtitle)`
      );

      const mapped: ClientNote[] = (result.entities ?? []).map((e: any) => {
        const ownerFull: string =
          e["ownerid@OData.Community.Display.V1.FormattedValue"] ??
          e["_ownerid_value@OData.Community.Display.V1.FormattedValue"] ??
          "Unknown";
        const statusVal: string = String(e["acc_ipfclientstatus"] ?? "");
        const statusLabel: string =
          e["acc_ipfclientstatus@OData.Community.Display.V1.FormattedValue"] ??
          statusVal;
        const relatedToLabel: string =
          e["acc_relatedto@OData.Community.Display.V1.FormattedValue"] ??
          e["acc_relatedto"] ??
          "";

        return {
          id: e["acc_clientnoteid"],
          acc_name: e["acc_name"] ?? "(No Title)",
          acc_ipfclientstatus: statusVal,
          acc_ipfclientstatusLabel: statusLabel,
          ownerName: ownerFull,
          ownerInitials: initials(ownerFull),
          ownerColor: avatarColor(ownerFull),
          position: e["ownerid"]?.jobtitle ?? "SPO",
          description: e["description"] ?? "",
          acc_relatedto: e["acc_relatedto"] ?? "",
          acc_relatedtoLabel: relatedToLabel,
          createdon: e["createdon"] ?? "",
        };
      });

      setNotes(mapped);
      setFiltered(mapped);
    } catch (err: any) {
      setError(err?.message ?? "Failed to load client notes.");
    } finally {
      setLoading(false);
    }
  }, [clientId, webAPI]);

  React.useEffect(() => {
    fetchNotes();
  }, [fetchNotes]);

  /* ── Search filter ─────────────────────── */
  React.useEffect(() => {
    const q = search.toLowerCase();
    if (!q) {
      setFiltered(notes);
    } else {
      setFiltered(
        notes.filter(
          (n) =>
            n.acc_name.toLowerCase().includes(q) ||
            n.description.toLowerCase().includes(q) ||
            n.ownerName.toLowerCase().includes(q) ||
            n.acc_ipfclientstatusLabel.toLowerCase().includes(q)
        )
      );
    }
  }, [search, notes]);

  /* ── Open record ───────────────────────── */
  const openRecord = React.useCallback((id: string) => {
    navigation.openForm({
      entityName: "acc_clientnote",
      entityId: id,
    });
  }, [navigation]);

  /* ── Create new ────────────────────────── */
  const createNew = React.useCallback(() => {
    navigation.openForm(
      { entityName: "acc_clientnote" },
      { acc_clientid: clientId }
    );
  }, [navigation, clientId]);

  /* ─────────────────────────────────────────
     Render
  ───────────────────────────────────────── */
  return (
    <div style={styles.wrapper}>
      {/* ── Toolbar ── */}
      <div style={styles.toolbar}>
        <span style={styles.title}>Timeline</span>
        <div style={styles.toolbarRight}>
          <button style={styles.btnPrimary} onClick={createNew}>
            <span style={styles.btnIcon}>+</span> New Client Note
          </button>
          <button style={styles.btnSecondary} onClick={fetchNotes}>
            <span style={styles.refreshIcon}>↻</span> Refresh
          </button>
        </div>
      </div>

      {/* ── Search row ── */}
      <div style={styles.filterRow}>
        <div style={styles.searchWrapper}>
          <span style={styles.searchIcon}>🔍</span>
          <input
            style={styles.searchInput}
            placeholder="Search Timeline"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
      </div>

      {/* ── Body ── */}
      {loading && <div style={styles.stateMsg}>Loading client notes…</div>}
      {error && <div style={{ ...styles.stateMsg, color: "#e74c3c" }}>{error}</div>}

      {!loading && !error && filtered.length === 0 && (
        <div style={styles.stateMsg}>No client notes found.</div>
      )}

      {!loading && !error && filtered.length > 0 && (
        <div style={styles.timeline}>
          {filtered.map((note, idx) => (
            <TimelineCard
              key={note.id}
              note={note}
              isLast={idx === filtered.length - 1}
              onOpen={openRecord}
              badgeColor={getStatusColor(note.acc_ipfclientstatus)}
            />
          ))}
        </div>
      )}
    </div>
  );
};

/* ─────────────────────────────────────────────
   Timeline Card
───────────────────────────────────────────── */
interface CardProps {
  note: ClientNote;
  isLast: boolean;
  onOpen: (id: string) => void;
  badgeColor: string;
}

const TimelineCard: React.FC<CardProps> = ({ note, isLast, onOpen, badgeColor }) => {
  const [hovered, setHovered] = React.useState(false);

  return (
    <div style={styles.cardRow}>
      <div style={styles.avatarCol}>
        <div style={{ ...styles.avatar, backgroundColor: note.ownerColor }}>
          {note.ownerInitials}
        </div>
        {!isLast && <div style={styles.connector} />}
      </div>

      <div
        style={{
          ...styles.card,
          boxShadow: hovered
            ? "0 2px 12px rgba(0,0,0,0.12)"
            : "0 1px 4px rgba(0,0,0,0.06)",
        }}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
      >
        <div style={styles.cardInner}>
          <div style={styles.cardLeft}>
            <div style={styles.cardHeader}>
              <span style={styles.noteTitle}>{note.acc_name}</span>
              <span style={styles.noteDate}>{fmtDate(note.createdon)}</span>
            </div>
            <div style={styles.byLine}>
              Created by {note.ownerName}, {note.position}
            </div>
            <p style={styles.description}>{note.description}</p>
          </div>

          <div style={styles.cardRight}>
            <div style={styles.relatedLabel}>Related to:</div>
            {note.acc_relatedtoLabel ? (
              <button
                style={{ ...styles.statusBadge, backgroundColor: badgeColor }}
                onClick={() => onOpen(note.id)}
                title="Open Client Note"
              >
                {note.acc_relatedtoLabel}
              </button>
            ) : (
              <span style={styles.noRelated}>—</span>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

/* ─────────────────────────────────────────────
   Styles
───────────────────────────────────────────── */
const styles: Record<string, React.CSSProperties> = {
  wrapper: {
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    color: "#201f1e",
    backgroundColor: "#ffffff",
    display: "flex",
    flexDirection: "column",
    height: "100%",
    minHeight: 300,
  },
  toolbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 16px",
    borderBottom: "1px solid #edebe9",
  },
  title: {
    fontWeight: 600,
    fontSize: 16,
    color: "#201f1e",
  },
  toolbarRight: {
    display: "flex",
    gap: 8,
    alignItems: "center",
  },
  btnPrimary: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    padding: "6px 14px",
    backgroundColor: "#ffffff",
    border: "2px solid #d63384",
    borderRadius: 4,
    color: "#201f1e",
    fontWeight: 600,
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  btnIcon: {
    fontSize: 16,
    color: "#d63384",
    fontWeight: 700,
    lineHeight: 1,
  },
  btnSecondary: {
    display: "flex",
    alignItems: "center",
    gap: 4,
    padding: "6px 12px",
    backgroundColor: "#ffffff",
    border: "1px solid #d2d0ce",
    borderRadius: 4,
    color: "#201f1e",
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  refreshIcon: { fontSize: 15 },
  filterRow: {
    display: "flex",
    alignItems: "center",
    padding: "8px 16px",
  },
  searchWrapper: {
    display: "flex",
    alignItems: "center",
    border: "1px solid #d2d0ce",
    borderRadius: 4,
    padding: "4px 8px",
    gap: 6,
    backgroundColor: "#faf9f8",
    width: 220,
  },
  searchIcon: { fontSize: 13, color: "#a19f9d" },
  searchInput: {
    border: "none",
    outline: "none",
    background: "transparent",
    fontSize: 13,
    color: "#201f1e",
    width: "100%",
  },
  timeline: {
    flex: 1,
    overflowY: "auto",
    padding: "8px 16px 16px",
  },
  cardRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: 12,
    marginBottom: 0,
  },
  avatarCol: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    flexShrink: 0,
    paddingTop: 16,
  },
  avatar: {
    width: 36,
    height: 36,
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#fff",
    fontWeight: 700,
    fontSize: 13,
    flexShrink: 0,
    zIndex: 1,
  },
  connector: {
    width: 2,
    flex: 1,
    minHeight: 24,
    backgroundColor: "#e1dfdd",
    marginTop: 0,
  },
  card: {
    flex: 1,
    border: "1px solid #edebe9",
    borderRadius: 6,
    padding: "14px 16px",
    backgroundColor: "#ffffff",
    transition: "box-shadow 0.15s ease",
    marginBottom: 16,
    marginTop: 8,
  },
  cardInner: {
    display: "flex",
    gap: 16,
    alignItems: "flex-start",
  },
  cardLeft: { flex: 1, minWidth: 0 },
  cardHeader: {
    display: "flex",
    alignItems: "baseline",
    gap: 10,
    marginBottom: 2,
  },
  noteTitle: { fontWeight: 700, fontSize: 14, color: "#201f1e" },
  noteDate: { fontSize: 12, color: "#605e5c" },
  byLine: { fontSize: 12, color: "#605e5c", marginBottom: 6 },
  description: {
    fontSize: 13,
    color: "#323130",
    lineHeight: 1.5,
    margin: 0,
    display: "-webkit-box",
    WebkitLineClamp: 3,
    WebkitBoxOrient: "vertical",
    overflow: "hidden",
  } as React.CSSProperties,
  cardRight: {
    flexShrink: 0,
    width: 120,
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-end",
    gap: 6,
  },
  relatedLabel: {
    fontSize: 12,
    fontWeight: 600,
    color: "#201f1e",
    textAlign: "right",
  },
  statusBadge: {
    padding: "6px 10px",
    borderRadius: 4,
    color: "#ffffff",
    fontSize: 12,
    fontWeight: 600,
    textAlign: "center",
    border: "none",
    cursor: "pointer",
    width: "100%",
    lineHeight: 1.4,
    wordBreak: "break-word",
    transition: "opacity 0.15s ease",
  },
  noRelated: { fontSize: 12, color: "#a19f9d" },
  stateMsg: {
    padding: "32px 16px",
    textAlign: "center",
    color: "#605e5c",
    fontSize: 13,
  },
};

  id: string;
  acc_name: string;
  acc_ipfclientstatus: string;
  acc_ipfclientstatusLabel: string;
  ownerName: string;
  ownerInitials: string;
  ownerColor: string;
  position: string; // SPO / job title
  description: string;
  acc_relatedto: string;
  acc_relatedtoLabel: string;
  createdon: string;
}

interface Props {
  clientId: string;
  webAPI: ComponentFramework.WebApi;
  navigation: ComponentFramework.Navigation;
  utils: ComponentFramework.Utility;
  containerWidth: number;
}

/* ─────────────────────────────────────────────
   Colour palette for acc_ipfclientstatus badges
   Extend / adjust option-set values as needed
───────────────────────────────────────────── */
const STATUS_COLOURS: Record<string, string> = {
  "1": "#e74c3c", // red
  "2": "#e67e22", // orange
  "3": "#f39c12", // yellow-orange  (3. Planning / CMP)
  "4": "#27ae60", // green           (4. Manage Change)
  "5": "#8e44ad", // purple          (5. Priorities)
  "6": "#2980b9", // blue
  "7": "#16a085", // teal
  default: "#95a5a6",
};

function statusColor(value: string): string {
  return STATUS_COLOURS[value] ?? STATUS_COLOURS["default"];
}

/* ─────────────────────────────────────────────
   Avatar colours keyed on initials hash
───────────────────────────────────────────── */
const AVATAR_COLOURS = [
  "#3498db", "#2ecc71", "#e74c3c", "#9b59b6",
  "#f39c12", "#1abc9c", "#e67e22", "#34495e",
];

function avatarColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash += name.charCodeAt(i);
  return AVATAR_COLOURS[hash % AVATAR_COLOURS.length];
}

function initials(name: string): string {
  return name
    .split(" ")
    .slice(0, 2)
    .map((n) => n[0]?.toUpperCase() ?? "")
    .join("");
}

/* ─────────────────────────────────────────────
   Date formatter  dd/MM/yyyy
───────────────────────────────────────────── */
function fmtDate(iso: string): string {
  if (!iso) return "";
  const d = new Date(iso);
  return `${String(d.getDate()).padStart(2, "0")}/${String(
    d.getMonth() + 1
  ).padStart(2, "0")}/${d.getFullYear()}`;
}

/* ─────────────────────────────────────────────
   Main component
───────────────────────────────────────────── */
export const ClientNoteTimelineComponent: React.FC<Props> = ({
  clientId,
  webAPI,
  navigation,
  utils,
  containerWidth,
}) => {
  const [notes, setNotes] = React.useState<ClientNote[]>([]);
  const [filtered, setFiltered] = React.useState<ClientNote[]>([]);
  const [search, setSearch] = React.useState("");
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  /* ── Fetch ─────────────────────────────── */
  const fetchNotes = React.useCallback(async () => {
    if (!clientId) return;
    setLoading(true);
    setError(null);
    try {
      const select = [
        "acc_clientnoteid",
        "acc_name",
        "acc_ipfclientstatus",
        "description",
        "createdon",
        "_ownerid_value",
        "acc_relatedto",
      ].join(",");

      const filter = `acc_clientid eq '${clientId}'`;
      const orderby = "createdon desc";

      const result = await webAPI.retrieveMultipleRecords(
        "acc_clientnote",
        `?$select=${select}&$filter=${filter}&$orderby=${orderby}&$expand=ownerid($select=fullname,jobtitle)`
      );

      const mapped: ClientNote[] = (result.entities ?? []).map((e: any) => {
        const ownerFull: string =
          e["ownerid@OData.Community.Display.V1.FormattedValue"] ??
          e["_ownerid_value@OData.Community.Display.V1.FormattedValue"] ??
          "Unknown";
        const statusVal: string = String(e["acc_ipfclientstatus"] ?? "");
        const statusLabel: string =
          e["acc_ipfclientstatus@OData.Community.Display.V1.FormattedValue"] ??
          statusVal;
        const relatedToLabel: string =
          e["acc_relatedto@OData.Community.Display.V1.FormattedValue"] ??
          e["acc_relatedto"] ??
          "";

        return {
          id: e["acc_clientnoteid"],
          acc_name: e["acc_name"] ?? "(No Title)",
          acc_ipfclientstatus: statusVal,
          acc_ipfclientstatusLabel: statusLabel,
          ownerName: ownerFull,
          ownerInitials: initials(ownerFull),
          ownerColor: avatarColor(ownerFull),
          position: e["ownerid"]?.jobtitle ?? "SPO",
          description: e["description"] ?? "",
          acc_relatedto: e["acc_relatedto"] ?? "",
          acc_relatedtoLabel: relatedToLabel,
          createdon: e["createdon"] ?? "",
        };
      });

      setNotes(mapped);
      setFiltered(mapped);
    } catch (err: any) {
      setError(err?.message ?? "Failed to load client notes.");
    } finally {
      setLoading(false);
    }
  }, [clientId, webAPI]);

  React.useEffect(() => {
    fetchNotes();
  }, [fetchNotes]);

  /* ── Search filter ─────────────────────── */
  React.useEffect(() => {
    const q = search.toLowerCase();
    if (!q) {
      setFiltered(notes);
    } else {
      setFiltered(
        notes.filter(
          (n) =>
            n.acc_name.toLowerCase().includes(q) ||
            n.description.toLowerCase().includes(q) ||
            n.ownerName.toLowerCase().includes(q) ||
            n.acc_ipfclientstatusLabel.toLowerCase().includes(q)
        )
      );
    }
  }, [search, notes]);

  /* ── Open record ───────────────────────── */
  const openRecord = (id: string) => {
    navigation.openForm({
      entityName: "acc_clientnote",
      entityId: id,
    });
  };

  /* ── Create new ────────────────────────── */
  const createNew = () => {
    navigation.openForm(
      {
        entityName: "acc_clientnote",
      },
      {
        acc_clientid: clientId,
      }
    );
  };

  /* ─────────────────────────────────────────
     Render
  ───────────────────────────────────────── */
  return (
    <div style={styles.wrapper}>
      {/* ── Toolbar ── */}
      <div style={styles.toolbar}>
        <span style={styles.title}>Timeline</span>
        <div style={styles.toolbarRight}>
          <button style={styles.btnPrimary} onClick={createNew}>
            <span style={styles.btnIcon}>+</span> New Client Note
          </button>
          <button style={styles.btnSecondary} onClick={fetchNotes}>
            <span style={styles.refreshIcon}>↻</span> Refresh
          </button>
        </div>
      </div>

      {/* ── Search + Related filter row ── */}
      <div style={styles.filterRow}>
        <div style={styles.searchWrapper}>
          <span style={styles.searchIcon}>🔍</span>
          <input
            style={styles.searchInput}
            placeholder="Search Timeline"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
      </div>

      {/* ── Body ── */}
      {loading && <div style={styles.stateMsg}>Loading client notes…</div>}
      {error && <div style={{ ...styles.stateMsg, color: "#e74c3c" }}>{error}</div>}

      {!loading && !error && filtered.length === 0 && (
        <div style={styles.stateMsg}>No client notes found.</div>
      )}

      {!loading && !error && filtered.length > 0 && (
        <div style={styles.timeline}>
          {filtered.map((note, idx) => (
            <TimelineCard
              key={note.id}
              note={note}
              isLast={idx === filtered.length - 1}
              onOpen={openRecord}
            />
          ))}
        </div>
      )}
    </div>
  );
};

/* ─────────────────────────────────────────────
   Timeline Card
───────────────────────────────────────────── */
interface CardProps {
  note: ClientNote;
  isLast: boolean;
  onOpen: (id: string) => void;
}

const TimelineCard: React.FC<CardProps> = ({ note, isLast, onOpen }) => {
  const [hovered, setHovered] = React.useState(false);

  return (
    <div style={styles.cardRow}>
      {/* Left: avatar + vertical line */}
      <div style={styles.avatarCol}>
        <div
          style={{
            ...styles.avatar,
            backgroundColor: note.ownerColor,
          }}
        >
          {note.ownerInitials}
        </div>
        {!isLast && <div style={styles.connector} />}
      </div>

      {/* Card body */}
      <div
        style={{
          ...styles.card,
          boxShadow: hovered
            ? "0 2px 12px rgba(0,0,0,0.12)"
            : "0 1px 4px rgba(0,0,0,0.06)",
        }}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
      >
        {/* Card top: header + date left, related-to right */}
        <div style={styles.cardInner}>
          <div style={styles.cardLeft}>
            <div style={styles.cardHeader}>
              <span style={styles.noteTitle}>{note.acc_name}</span>
              <span style={styles.noteDate}>{fmtDate(note.createdon)}</span>
            </div>
            <div style={styles.byLine}>
              Created by {note.ownerName}, {note.position}
            </div>
            <p style={styles.description}>{note.description}</p>
          </div>

          {/* Related-to + coloured status badge */}
          <div style={styles.cardRight}>
            <div style={styles.relatedLabel}>Related to:</div>
            {note.acc_relatedtoLabel ? (
              <button
                style={{
                  ...styles.statusBadge,
                  backgroundColor: statusColor(note.acc_ipfclientstatus),
                }}
                onClick={() => onOpen(note.id)}
                title="Open Client Note"
              >
                {note.acc_relatedtoLabel}
              </button>
            ) : (
              <span style={styles.noRelated}>—</span>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

/* ─────────────────────────────────────────────
   Styles
───────────────────────────────────────────── */
const styles: Record<string, React.CSSProperties> = {
  wrapper: {
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    color: "#201f1e",
    backgroundColor: "#ffffff",
    display: "flex",
    flexDirection: "column",
    height: "100%",
    minHeight: 300,
  },

  /* Toolbar */
  toolbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 16px",
    borderBottom: "1px solid #edebe9",
  },
  title: {
    fontWeight: 600,
    fontSize: 16,
    color: "#201f1e",
  },
  toolbarRight: {
    display: "flex",
    gap: 8,
    alignItems: "center",
  },
  btnPrimary: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    padding: "6px 14px",
    backgroundColor: "#ffffff",
    border: "2px solid #d63384",
    borderRadius: 4,
    color: "#201f1e",
    fontWeight: 600,
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  btnIcon: {
    fontSize: 16,
    color: "#d63384",
    fontWeight: 700,
    lineHeight: 1,
  },
  btnSecondary: {
    display: "flex",
    alignItems: "center",
    gap: 4,
    padding: "6px 12px",
    backgroundColor: "#ffffff",
    border: "1px solid #d2d0ce",
    borderRadius: 4,
    color: "#201f1e",
    fontSize: 13,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
  refreshIcon: {
    fontSize: 15,
  },

  /* Filter row */
  filterRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "8px 16px",
  },
  searchWrapper: {
    display: "flex",
    alignItems: "center",
    border: "1px solid #d2d0ce",
    borderRadius: 4,
    padding: "4px 8px",
    gap: 6,
    backgroundColor: "#faf9f8",
    width: 220,
  },
  searchIcon: {
    fontSize: 13,
    color: "#a19f9d",
  },
  searchInput: {
    border: "none",
    outline: "none",
    background: "transparent",
    fontSize: 13,
    color: "#201f1e",
    width: "100%",
  },

  /* Timeline list */
  timeline: {
    flex: 1,
    overflowY: "auto",
    padding: "8px 16px 16px",
  },

  /* Card row (avatar + card) */
  cardRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: 12,
    marginBottom: 0,
  },

  /* Avatar column */
  avatarCol: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    flexShrink: 0,
    paddingTop: 16,
  },
  avatar: {
    width: 36,
    height: 36,
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#fff",
    fontWeight: 700,
    fontSize: 13,
    flexShrink: 0,
    zIndex: 1,
  },
  connector: {
    width: 2,
    flex: 1,
    minHeight: 24,
    backgroundColor: "#e1dfdd",
    marginTop: 0,
  },

  /* Card */
  card: {
    flex: 1,
    border: "1px solid #edebe9",
    borderRadius: 6,
    padding: "14px 16px",
    backgroundColor: "#ffffff",
    transition: "box-shadow 0.15s ease",
    marginBottom: 16,
    marginTop: 8,
  },
  cardInner: {
    display: "flex",
    gap: 16,
    alignItems: "flex-start",
  },
  cardLeft: {
    flex: 1,
    minWidth: 0,
  },
  cardHeader: {
    display: "flex",
    alignItems: "baseline",
    gap: 10,
    marginBottom: 2,
  },
  noteTitle: {
    fontWeight: 700,
    fontSize: 14,
    color: "#201f1e",
  },
  noteDate: {
    fontSize: 12,
    color: "#605e5c",
  },
  byLine: {
    fontSize: 12,
    color: "#605e5c",
    marginBottom: 6,
  },
  description: {
    fontSize: 13,
    color: "#323130",
    lineHeight: 1.5,
    margin: 0,
    display: "-webkit-box",
    WebkitLineClamp: 3,
    WebkitBoxOrient: "vertical",
    overflow: "hidden",
  } as React.CSSProperties,

  /* Right side - related to */
  cardRight: {
    flexShrink: 0,
    width: 120,
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-end",
    gap: 6,
  },
  relatedLabel: {
    fontSize: 12,
    fontWeight: 600,
    color: "#201f1e",
    textAlign: "right",
  },
  statusBadge: {
    padding: "6px 10px",
    borderRadius: 4,
    color: "#ffffff",
    fontSize: 12,
    fontWeight: 600,
    textAlign: "center",
    border: "none",
    cursor: "pointer",
    width: "100%",
    lineHeight: 1.4,
    wordBreak: "break-word",
    transition: "opacity 0.15s ease",
  },
  noRelated: {
    fontSize: 12,
    color: "#a19f9d",
  },

  /* State messages */
  stateMsg: {
    padding: "32px 16px",
    textAlign: "center",
    color: "#605e5c",
    fontSize: 13,
  },
};
