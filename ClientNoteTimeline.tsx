import * as React from "react";

/* ─────────────────────────────────────────────
   Types
───────────────────────────────────────────── */
export interface ClientNote {
  id: string;
  acc_name: string;
  acc_ipfclientstatus: string;
  acc_ipfclientstatusLabel: string;
  ownerId: string;
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
  dataset: ComponentFramework.PropertyTypes.DataSet;
  clientId: string;
  webAPI: ComponentFramework.WebApi;
  navigation: ComponentFramework.Navigation;
  utils: ComponentFramework.Utility;
  containerWidth: number;
  containerHeight: number;
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
   Read dataset rows into ClientNote[]
   The dataset is already filtered by the subgrid
   binding — no manual OData filter needed.
───────────────────────────────────────────── */
function mapDataset(
  dataset: ComponentFramework.PropertyTypes.DataSet
): ClientNote[] {
  return dataset.sortedRecordIds.map((id) => {
    const rec = dataset.records[id];

    const getStr = (col: string): string => {
      try { return rec.getFormattedValue(col) ?? ""; } catch { return ""; }
    };
    const getRaw = (col: string): string => {
      try {
        const v = rec.getValue(col);
        return v != null ? String(v) : "";
      } catch { return ""; }
    };

    const ownerFull = getStr("ownerid") || "Unknown";
    const statusVal = getRaw("acc_ipfclientstatus");
    const statusLabel = getStr("acc_ipfclientstatus") || statusVal;
    const relatedToLabel = getStr("acc_relatedto");

    // Extract the raw owner GUID for the position lookup
    const ownerRef = rec.getValue("ownerid") as any;
    const ownerId: string =
      ownerRef?.id?.guid ??
      ownerRef?.id ??
      getRaw("_ownerid_value") ??
      "";

    return {
      id,
      acc_name: getStr("acc_name") || "(No Title)",
      acc_ipfclientstatus: statusVal,
      acc_ipfclientstatusLabel: statusLabel,
      ownerId,
      ownerName: ownerFull,
      ownerInitials: initials(ownerFull),
      ownerColor: avatarColor(ownerFull),
      position: "", // resolved async via fetchOwnerPositions
      description: getStr("description"),
      acc_relatedto: getRaw("acc_relatedto"),
      acc_relatedtoLabel: relatedToLabel,
      createdon: getRaw("createdon"),
    };
  });
}

/* ─────────────────────────────────────────────
   Fetch option set colours from metadata
   Returns { optionValue -> hexColor }
───────────────────────────────────────────── */
async function fetchOptionSetColors(
  webAPI: ComponentFramework.WebApi
): Promise<Record<string, string>> {
  try {
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
      const raw = opt.Color ?? opt.color ?? null;
      if (raw !== null && raw !== undefined) {
        map[value] =
          typeof raw === "number"
            ? "#" + ((raw & 0xffffff) >>> 0).toString(16).padStart(6, "0")
            : String(raw).startsWith("#")
            ? raw
            : `#${raw}`;
      }
    }
    return map;
  } catch {
    return {};
  }
}


/* ─────────────────────────────────────────────
   Fetch owner job titles in a single batch call
   Returns { systemuserid -> jobtitle }
───────────────────────────────────────────── */
async function fetchOwnerPositions(
  webAPI: ComponentFramework.WebApi,
  ownerIds: string[]
): Promise<Record<string, string>> {
  if (ownerIds.length === 0) return {};
  try {
    const unique = [...new Set(ownerIds.filter(Boolean))];
    if (unique.length === 0) return {};
    const filter = unique
      .map((id) => `systemuserid eq '${id}'`)
      .join(" or ");
    const result = await webAPI.retrieveMultipleRecords(
      "systemuser",
      `?$select=systemuserid,jobtitle&$filter=${filter}`
    );
    const map: Record<string, string> = {};
    for (const u of result.entities ?? []) {
      map[u["systemuserid"]] = u["jobtitle"] ?? "";
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
  dataset,
  clientId,
  webAPI,
  navigation,
  containerWidth,
  containerHeight,
}) => {
  const [notes, setNotes] = React.useState<ClientNote[]>([]);
  const [filtered, setFiltered] = React.useState<ClientNote[]>([]);
  const [search, setSearch] = React.useState("");
  const [statusColorMap, setStatusColorMap] = React.useState<Record<string, string>>({});
  const [positionMap, setPositionMap] = React.useState<Record<string, string>>({});

  /* ── Map dataset rows whenever dataset updates ── */
  React.useEffect(() => {
    if (!dataset.loading) {
      const mapped = mapDataset(dataset);
      setNotes(mapped);
      setFiltered(mapped);

      // Fetch job titles for all unique owners in this dataset
      const ownerIds = [...new Set(mapped.map((n) => n.ownerId).filter(Boolean))];
      fetchOwnerPositions(webAPI, ownerIds).then(setPositionMap);
    }
  }, [dataset, dataset.loading, dataset.sortedRecordIds]);

  /* ── Fetch option set colours once ── */
  React.useEffect(() => {
    fetchOptionSetColors(webAPI).then(setStatusColorMap);
  }, [webAPI]);

  /* ── Resolve badge colour ── */
  const getStatusColor = React.useCallback(
    (value: string): string => statusColorMap[value] ?? DEFAULT_BADGE_COLOR,
    [statusColorMap]
  );

  /* ── Search filter ── */
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

  /* ── Refresh — tell the dataset to reload ── */
  const handleRefresh = React.useCallback(() => {
    dataset.refresh();
  }, [dataset]);

  /* ── Open existing record ── */
  const openRecord = React.useCallback(
    (id: string) => {
      navigation.openForm({
        entityName: "acc_clientnote",
        entityId: id,
      });
    },
    [navigation]
  );

  /* ── Create new record — clientId passed as default ── */
  const createNew = React.useCallback(() => {
    navigation.openForm(
      { entityName: "acc_clientnote" },
      { acc_clientid: clientId }
    );
  }, [navigation, clientId]);

  const isLoading = dataset.loading;

  /* ─────────────────────────────────────────
     Render — wrapper fills allocated subgrid
     width and height from PCF container
  ───────────────────────────────────────── */
  return (
    <div
      style={{
        ...styles.wrapper,
        width: containerWidth > 0 ? containerWidth : "100%",
        minHeight: containerHeight > 0 ? containerHeight : 400,
      }}
    >
      {/* ── Toolbar ── */}
      <div style={styles.toolbar}>
        <span style={styles.title}>Timeline</span>
        <div style={styles.toolbarRight}>
          <button style={styles.btnPrimary} onClick={createNew}>
            <span style={styles.btnIcon}>+</span> New Client Note
          </button>
          <button style={styles.btnSecondary} onClick={handleRefresh}>
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
      {isLoading && <div style={styles.stateMsg}>Loading client notes…</div>}

      {!isLoading && filtered.length === 0 && (
        <div style={styles.stateMsg}>No client notes found.</div>
      )}

      {!isLoading && filtered.length > 0 && (
        <div style={styles.timeline}>
          {filtered.map((note, idx) => (
            <TimelineCard
              key={note.id}
              note={{ ...note, position: positionMap[note.ownerId] || note.position || "SPO" }}
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
      {/* Avatar + connector */}
      <div style={styles.avatarCol}>
        <div style={{ ...styles.avatar, backgroundColor: note.ownerColor }}>
          {note.ownerInitials}
        </div>
        {!isLast && <div style={styles.connector} />}
      </div>

      {/* Card */}
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
          {/* Left: title, byline, description */}
          <div style={styles.cardLeft}>
            <div style={styles.cardHeader}>
              <span style={styles.noteTitle}>{note.acc_name}</span>
              <span style={styles.noteDate}>{fmtDate(note.createdon)}</span>
            </div>
            <div style={styles.byLine}>
              Created by {note.ownerName}, {note.position}
            </div>
            <p style={styles.description as React.CSSProperties}>{note.description}</p>
          </div>

          {/* Right: Related to + coloured badge */}
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
  /* Wrapper — fills full subgrid width */
  wrapper: {
    fontFamily: "'Segoe UI', system-ui, sans-serif",
    fontSize: 14,
    color: "#201f1e",
    backgroundColor: "#ffffff",
    display: "flex",
    flexDirection: "column",
    boxSizing: "border-box",
    overflow: "hidden",
  },

  toolbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 16px",
    borderBottom: "1px solid #edebe9",
    flexShrink: 0,
  },
  title: { fontWeight: 600, fontSize: 16, color: "#201f1e" },
  toolbarRight: { display: "flex", gap: 8, alignItems: "center" },

  btnPrimary: {
    display: "flex", alignItems: "center", gap: 6,
    padding: "6px 14px",
    backgroundColor: "#ffffff",
    border: "2px solid #d63384",
    borderRadius: 4,
    color: "#201f1e", fontWeight: 600, fontSize: 13,
    cursor: "pointer", whiteSpace: "nowrap",
  },
  btnIcon: { fontSize: 16, color: "#d63384", fontWeight: 700, lineHeight: 1 },

  btnSecondary: {
    display: "flex", alignItems: "center", gap: 4,
    padding: "6px 12px",
    backgroundColor: "#ffffff",
    border: "1px solid #d2d0ce",
    borderRadius: 4,
    color: "#201f1e", fontSize: 13,
    cursor: "pointer", whiteSpace: "nowrap",
  },
  refreshIcon: { fontSize: 15 },

  filterRow: {
    display: "flex", alignItems: "center",
    padding: "8px 16px",
    flexShrink: 0,
  },
  searchWrapper: {
    display: "flex", alignItems: "center",
    border: "1px solid #d2d0ce", borderRadius: 4,
    padding: "4px 8px", gap: 6,
    backgroundColor: "#faf9f8",
    width: 240,
  },
  searchIcon: { fontSize: 13, color: "#a19f9d" },
  searchInput: {
    border: "none", outline: "none",
    background: "transparent", fontSize: 13,
    color: "#201f1e", width: "100%",
  },

  /* Timeline scrollable area — fills remaining height */
  timeline: {
    flex: 1,
    overflowY: "auto",
    padding: "8px 16px 16px",
    width: "100%",
    boxSizing: "border-box",
  },

  cardRow: {
    display: "flex", alignItems: "flex-start",
    gap: 12, marginBottom: 0,
    width: "100%", boxSizing: "border-box",
  },

  avatarCol: {
    display: "flex", flexDirection: "column",
    alignItems: "center", flexShrink: 0,
    paddingTop: 16,
  },
  avatar: {
    width: 36, height: 36, borderRadius: "50%",
    display: "flex", alignItems: "center", justifyContent: "center",
    color: "#fff", fontWeight: 700, fontSize: 13,
    flexShrink: 0, zIndex: 1,
  },
  connector: {
    width: 2, flex: 1, minHeight: 24,
    backgroundColor: "#e1dfdd",
  },

  /* Card stretches to fill remaining row width */
  card: {
    flex: 1,
    minWidth: 0,
    border: "1px solid #edebe9",
    borderRadius: 6,
    padding: "14px 16px",
    backgroundColor: "#ffffff",
    transition: "box-shadow 0.15s ease",
    marginBottom: 16, marginTop: 8,
    boxSizing: "border-box",
  },

  cardInner: {
    display: "flex", gap: 16,
    alignItems: "flex-start",
    width: "100%",
  },
  cardLeft: { flex: 1, minWidth: 0 },

  cardHeader: {
    display: "flex", alignItems: "baseline",
    gap: 10, marginBottom: 2,
  },
  noteTitle: { fontWeight: 700, fontSize: 14, color: "#201f1e" },
  noteDate: { fontSize: 12, color: "#605e5c" },
  byLine: { fontSize: 12, color: "#605e5c", marginBottom: 6 },
  description: {
    fontSize: 13, color: "#323130",
    lineHeight: 1.5, margin: 0,
    display: "-webkit-box",
    WebkitLineClamp: 3,
    WebkitBoxOrient: "vertical",
    overflow: "hidden",
  },

  cardRight: {
    flexShrink: 0, width: 130,
    display: "flex", flexDirection: "column",
    alignItems: "flex-end", gap: 6,
  },
  relatedLabel: {
    fontSize: 12, fontWeight: 600,
    color: "#201f1e", textAlign: "right",
  },
  statusBadge: {
    padding: "6px 10px", borderRadius: 4,
    color: "#ffffff", fontSize: 12, fontWeight: 600,
    textAlign: "center", border: "none",
    cursor: "pointer", width: "100%",
    lineHeight: 1.4, wordBreak: "break-word",
    transition: "opacity 0.15s ease",
  },
  noRelated: { fontSize: 12, color: "#a19f9d" },

  stateMsg: {
    padding: "32px 16px",
    textAlign: "center",
    color: "#605e5c", fontSize: 13,
  },
};
