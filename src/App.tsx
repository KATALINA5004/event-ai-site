import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

type EventStatus = "upcoming" | "past" | "cancelled" | "draft" | "needs_review";
type Page = "all" | "date" | "category" | "mood" | "collections" | "favorites" | "about" | "admin";
type CollectionKey = "🔥 Лучшее и актуальное" | "💃 Тело и движение" | "👨‍👩‍👧 Семья и дети" | "💛 Отношения и развитие" | "🌿 Отдых и перезагрузка" | "🎲 Игры и досуг";

type EventItem = {
  id: string;
  title: string;
  shortDescription: string;
  description: string;
  date: string;
  time: string;
  city: string;
  location: string;
  address: string;
  price: string;
  isFree: boolean;
  category: string;
  categories: string[];
  mood: string;
  moods: string[];
  collections: string[];
  tags: string[];
  detailsLink: string;
  ticketLink: string;
  image: string;
  organizer: string;
  isVip: boolean;
  status: EventStatus;
  source: string;
  aiConfidence: number;
  createdAt: string;
  updatedAt: string;
};

type CollectionCard = {
  id: string;
  collections: CollectionKey[];
  title: string;
  image: string;
  description: string;
  link: string;
  createdAt: string;
};

type ProcessingReport = { loaded: number; added: number; updated: number; hiddenPast: number; cancelled: number; review: number; duplicates: number; published: number };
type AdminSubPage = "dashboard";

const categories = ["🎉 Тусовки", "🏃 Активности", "👨‍👩‍👧 Семья", "🌳 Природа", "🧘 Практики", "📚 Обучение", "🎮 Игры", "🎨 Искусство", "💼 Бизнес"];
const moods = ["🎉 Потусить", "✨ Вдохновиться", "🧘 Отдохнуть", "🎮 Поиграть", "🌈 Узнать новое", "🌿 На природу", "💛 Пообщаться", "📚 Развиваться"];
const collectionKeys: CollectionKey[] = ["🔥 Лучшее и актуальное", "💃 Тело и движение", "👨‍👩‍👧 Семья и дети", "💛 Отношения и развитие", "🌿 Отдых и перезагрузка", "🎲 Игры и досуг"];
const emptyReport: ProcessingReport = { loaded: 0, added: 0, updated: 0, hiddenPast: 0, cancelled: 0, review: 0, duplicates: 0, published: 0 };
const ADMIN_LOGIN = "admin";
const ADMIN_PASSWORD = "Atelier2026!";
const TELEGRAM_CHANNEL_URL = "https://t.me/testyar";
const TELEGRAM_CHAT_ID = "@testyar";
/** Личные сообщения организаторам */
const TELEGRAM_CONTACT_URL = "https://t.me/katalina5004";
const TELEGRAM_CONTACT_HANDLE = "@katalina5004";
const TELEGRAM_BOT_TOKEN = import.meta.env.VITE_TELEGRAM_BOT_TOKEN as string | undefined;

const normalize = (v: unknown) => String(v ?? "").trim();
const lower = (v: unknown) => normalize(v).toLowerCase();
const parse = <T,>(raw: string, fallback: T): T => { try { return JSON.parse(raw) as T; } catch { return fallback; } };
const short = (d: string) => (normalize(d).length > 160 ? `${normalize(d).slice(0, 157)}...` : normalize(d) || "Описание будет добавлено.");
const asArray = <T,>(value: unknown, fallback: T[]): T[] => (Array.isArray(value) ? (value as T[]) : fallback);
const asObject = <T extends object>(value: unknown, fallback: T): T =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as T) : fallback;

const stripLeadingEmoji = (value: string): string =>
  value
    .replace(
      // remove leading emoji / symbols and extra spaces (handles "🏃 Активности" and also "Активности")
      /^[\s\p{Extended_Pictographic}\p{Emoji_Presentation}\p{Emoji}\u200d\uFE0F]+/gu,
      ""
    )
    .replace(/^[\s\-–—•·]+/g, "")
    .trim();

const detectCategory = (text: string): string => {
  if (text.includes("вечерин") || text.includes("dj") || text.includes("бар")) return "🎉 Тусовки";
  if (text.includes("бег") || text.includes("спорт")) return "🏃 Активности";
  if (text.includes("сем") || text.includes("дет")) return "👨‍👩‍👧 Семья";
  if (text.includes("природ") || text.includes("лес") || text.includes("парк")) return "🌳 Природа";
  if (text.includes("йога") || text.includes("медита")) return "🧘 Практики";
  if (text.includes("лекц") || text.includes("мастер")) return "📚 Обучение";
  if (text.includes("квиз") || text.includes("игр")) return "🎮 Игры";
  if (text.includes("выстав") || text.includes("искус")) return "🎨 Искусство";
  return "💼 Бизнес";
};

const detectMood = (text: string): string => {
  if (text.includes("вечерин") || text.includes("концерт")) return "🎉 Потусить";
  if (text.includes("арт") || text.includes("театр")) return "✨ Вдохновиться";
  if (text.includes("йога") || text.includes("релакс")) return "🧘 Отдохнуть";
  if (text.includes("квиз") || text.includes("игр")) return "🎮 Поиграть";
  if (text.includes("лекц") || text.includes("новое")) return "🌈 Узнать новое";
  if (text.includes("лес") || text.includes("парк")) return "🌿 На природу";
  if (text.includes("нетворк") || text.includes("общ")) return "💛 Пообщаться";
  return "📚 Развиваться";
};

const detectCollections = (category: string, mood: string): string[] => {
  const out = new Set<string>(["🔥 Лучшее и актуальное"]);
  if (category.includes("Семья")) out.add("👨‍👩‍👧 Семья и дети");
  if (category.includes("Игры")) out.add("🎲 Игры и досуг");
  if (mood.includes("Подвигаться") || category.includes("Актив")) out.add("💃 Тело и движение");
  if (mood.includes("Пообщаться")) out.add("💛 Отношения и развитие");
  if (mood.includes("Отдохнуть") || category.includes("Природа")) out.add("🌿 Отдых и перезагрузка");
  return [...out];
};

const tagsFromCategoriesAndMoods = (cats: string[], moodsList: string[]): string[] =>
  Array.from(new Set([...cats, ...moodsList].map((x) => normalize(x.replace(/^.. /, ""))).filter(Boolean)));

const parseDateTime = (e: EventItem): number => {
  const dt = new Date(`${e.date}T${e.time || "23:59"}`);
  return Number.isNaN(dt.getTime()) ? Number.MAX_SAFE_INTEGER : dt.getTime();
};

const weekdayShortRu = (iso: string): string => {
  const dt = new Date(`${iso}T12:00`);
  if (Number.isNaN(dt.getTime())) return "";
  return dt.toLocaleDateString("ru-RU", { weekday: "short" }).replace(/\.$/, "").trim();
};

const toIsoDate = (year: number, month: number, day: number): string => {
  const y = String(year).padStart(4, "0");
  const m = String(month).padStart(2, "0");
  const d = String(day).padStart(2, "0");
  return `${y}-${m}-${d}`;
};

const normalizeDateInput = (value: unknown): string => {
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return toIsoDate(parsed.y, parsed.m, parsed.d);
    return "";
  }
  const raw = normalize(value);
  if (!raw) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(raw)) {
    const [d, m, y] = raw.split(/[./-]/).map((x) => Number(x));
    return toIsoDate(y < 100 ? y + 2000 : y, m, d);
  }
  if (/^\d{5}$/.test(raw)) {
    const parsed = XLSX.SSF.parse_date_code(Number(raw));
    if (parsed) return toIsoDate(parsed.y, parsed.m, parsed.d);
  }
  const ruMatch = raw.match(/^(\d{1,2})\s+([а-яё]+)(?:\s+(\d{4}))?$/i);
  if (ruMatch) {
    const day = Number(ruMatch[1]);
    const monthName = ruMatch[2].toLowerCase();
    const year = Number(ruMatch[3] || new Date().getFullYear());
    const monthMap: Record<string, number> = {
      января: 1, феврал: 2, марта: 3, апрел: 4, мая: 5, июн: 6, июл: 7, август: 8, сентябр: 9, октябр: 10, ноябр: 11, декабр: 12
    };
    const monthEntry = Object.entries(monthMap).find(([k]) => monthName.startsWith(k));
    if (monthEntry) return toIsoDate(year, monthEntry[1], day);
  }
  const jsDate = new Date(raw);
  if (!Number.isNaN(jsDate.getTime())) return jsDate.toISOString().slice(0, 10);
  return "";
};

const normalizeTimeInput = (value: unknown): string => {
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${String(parsed.H).padStart(2, "0")}:${String(parsed.M).padStart(2, "0")}`;
  }
  const raw = normalize(value);
  if (!raw) return "";
  const rangeMatch = raw.match(/^(\d{1,2}:\d{2})\s*[-–]\s*\d{1,2}:\d{2}$/);
  if (rangeMatch) return rangeMatch[1].padStart(5, "0");
  if (/^\d{1,2}:\d{2}$/.test(raw)) return raw.padStart(5, "0");
  if (/^\d{1,2}\.\d{2}$/.test(raw)) return raw.replace(".", ":");
  if (/^0\.\d+$/.test(raw)) {
    const minutes = Math.round(Number(raw) * 24 * 60);
    const h = Math.floor(minutes / 60) % 24;
    const m = minutes % 60;
    return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
  }
  return "";
};

const bumpDateToFuture = (date: string, time: string): string => {
  if (!date) return "";
  const dt = new Date(`${date}T${time || "23:59"}`);
  if (Number.isNaN(dt.getTime())) return date;
  const now = Date.now();
  // Useful for recurring/event calendars where exported year is stale.
  while (dt.getTime() < now) {
    dt.setFullYear(dt.getFullYear() + 1);
  }
  return dt.toISOString().slice(0, 10);
};

const normalizeKey = (k: string): string => lower(k).replace(/[\s_\-:()]/g, "");

const SEMANTIC_FIELDS = [
  "date",
  "time",
  "title",
  "description",
  "shortDescription",
  "location",
  "city",
  "address",
  "price",
  "tags",
  "detailsLink",
  "ticketLink",
  "image",
  "organizer",
  "category",
  "mood",
  "status"
] as const;
type SemanticField = (typeof SEMANTIC_FIELDS)[number];

const FIELD_ROW_KEYS: Record<SemanticField, string> = {
  title: "title",
  date: "date",
  time: "time",
  city: "city",
  location: "location",
  address: "address",
  price: "price",
  description: "description",
  shortDescription: "shortDescription",
  detailsLink: "detailsLink",
  ticketLink: "ticketLink",
  tags: "tags",
  image: "image",
  organizer: "organizer",
  category: "category",
  mood: "mood",
  status: "status"
};

const ASSIGN_ORDER: SemanticField[] = [
  "date",
  "time",
  "title",
  "description",
  "shortDescription",
  "location",
  "city",
  "address",
  "price",
  "tags",
  "detailsLink",
  "ticketLink",
  "image",
  "organizer",
  "category",
  "mood",
  "status"
];

function headerHintScore(headerNorm: string, field: SemanticField): number {
  const hints: Record<SemanticField, string[]> = {
    title: ["назван", "title", "name", "мероприят", "event", "заголовок", "тема"],
    date: ["дата", "date", "день", "когда"],
    time: ["время", "time", "час", "начало"],
    city: ["город", "city"],
    location: ["место", "локац", "venue", "площадк", "location"],
    address: ["адрес", "address"],
    price: ["цена", "price", "стоим", "билет"],
    description: ["описан", "description", "about", "подробн"],
    shortDescription: ["кратк", "short"],
    detailsLink: ["ссылк", "link", "url", "сайт", "подробнссыл"],
    ticketLink: ["билет", "регистрац", "ticket"],
    tags: ["тег", "tag", "ключ", "метк"],
    image: ["фото", "image", "картинк", "изображ"],
    organizer: ["организ"],
    category: ["категор"],
    mood: ["настроен", "mood"],
    status: ["статус", "status"]
  };
  return hints[field].some((h) => headerNorm.includes(normalizeKey(h))) ? 14 : 0;
}

function cellContentScore(cell: unknown, field: SemanticField): number {
  const s = normalize(cell);
  if (!s) return 0;
  switch (field) {
    case "date":
      return normalizeDateInput(cell) ? 10 : 0;
    case "time":
      return normalizeTimeInput(cell) ? 10 : 0;
    case "detailsLink":
    case "ticketLink":
    case "image":
      return /^https?:\/\//i.test(s) ? 10 : 0;
    case "price":
      return (/(\d|бесплат|free|thb|rub|usd|€|\$)/i.test(s) ? 5 : 0) + (/\d/.test(s) ? 4 : 0);
    case "tags":
      return /[,;|]/.test(s) && s.length < 200 ? 7 : 0;
    case "title":
      return s.length >= 6 && s.length <= 220 && !/^https?:/i.test(s) ? 5 : 0;
    case "description":
      return s.length > 35 ? 6 : s.length > 12 ? 3 : 0;
    case "shortDescription":
      return s.length > 8 && s.length < 120 ? 4 : 0;
    case "city":
    case "location":
    case "address":
      return s.length >= 2 && s.length < 120 && !/^https?:/i.test(s) ? 3 : 0;
    case "organizer":
      return s.length >= 2 && s.length < 80 ? 3 : 0;
    case "category":
    case "mood":
    case "status":
      return s.length ? 2 : 0;
    default:
      return 0;
  }
}

function isHeaderLikeRow(row: unknown[]): boolean {
  let scored = 0;
  let labelLike = 0;
  for (const cell of row) {
    const s = normalize(cell);
    if (!s) continue;
    scored++;
    if (normalizeDateInput(cell) || normalizeTimeInput(cell)) continue;
    if (/^https?:\/\//i.test(s)) continue;
    if (s.length <= 52) labelLike++;
  }
  return scored > 0 && labelLike / scored >= 0.52;
}

function inferColumnScores(aoa: unknown[][], headerRow: number, numCols: number): Record<SemanticField, number[]> {
  const headerNorms = Array.from({ length: numCols }, (_, i) =>
    headerRow >= 0 ? normalizeKey(String(aoa[headerRow]?.[i] ?? "")) : ""
  );
  const dataStart = headerRow < 0 ? 0 : headerRow + 1;
  const samples = aoa.slice(dataStart, dataStart + 120);

  const scores = Object.fromEntries(SEMANTIC_FIELDS.map((f) => [f, Array(numCols).fill(0)])) as Record<SemanticField, number[]>;

  for (let c = 0; c < numCols; c++) {
    const hh = headerNorms[c];
    for (const f of SEMANTIC_FIELDS) {
      let sc = hh ? headerHintScore(hh, f) : 0;
      let nonEmpty = 0;
      for (const row of samples) {
        const cell = row[c];
        if (!normalize(cell)) continue;
        nonEmpty++;
        sc += cellContentScore(cell, f);
      }
      if (nonEmpty) sc += Math.min(8, nonEmpty / 3);
      scores[f][c] = sc;
    }
  }
  return scores;
}

function greedyAssignScores(scores: Record<SemanticField, number[]>): Partial<Record<SemanticField, number>> {
  const numCols = scores.title?.length ?? 0;
  const used = new Set<number>();
  const mapping: Partial<Record<SemanticField, number>> = {};
  for (const f of ASSIGN_ORDER) {
    let bestC = -1;
    let bestS = -1;
    for (let c = 0; c < numCols; c++) {
      if (used.has(c)) continue;
      const s = scores[f][c];
      if (s > bestS) {
        bestS = s;
        bestC = c;
      }
    }
    if (bestC >= 0 && bestS >= 5) {
      mapping[f] = bestC;
      used.add(bestC);
    }
  }
  return mapping;
}

function aoaToSemanticRows(aoa: unknown[][]): Record<string, unknown>[] {
  const widths = aoa.map((r) => r.length);
  const numCols = widths.length ? Math.max(...widths) : 0;
  if (!numCols) return [];

  const headerRow = aoa.length && isHeaderLikeRow(aoa[0]) ? 0 : -1;
  const dataStart = headerRow < 0 ? 0 : headerRow + 1;
  const scores = inferColumnScores(aoa, headerRow, numCols);
  const mapping = greedyAssignScores(scores);
  const usedCols = new Set<number>(Object.values(mapping).filter((x): x is number => typeof x === "number"));

  const rows: Record<string, unknown>[] = [];
  for (let r = dataStart; r < aoa.length; r++) {
    const row = aoa[r] ?? [];
    if (!row.some((cell) => normalize(cell))) continue;

    const obj: Record<string, unknown> = {};
    for (const f of SEMANTIC_FIELDS) {
      const ci = mapping[f];
      if (ci === undefined) continue;
      const raw = row[ci];
      if (!normalize(raw)) continue;
      const key = FIELD_ROW_KEYS[f] as string;
      obj[key] = raw;
    }

    const extras: string[] = [];
    for (let c = 0; c < numCols; c++) {
      if (usedCols.has(c)) continue;
      const v = normalize(row[c]);
      if (v) extras.push(v);
    }
    if (extras.length) obj.__extras = extras.join("\n");

    rows.push(obj);
  }
  return rows;
}

const normalizeCategoryInput = (value: string): string => {
  const v = lower(value);
  if (!v) return "";
  if (v.includes("тусов")) return "🎉 Тусовки";
  if (v.includes("актив")) return "🏃 Активности";
  if (v.includes("сем")) return "👨‍👩‍👧 Семья";
  if (v.includes("природ")) return "🌳 Природа";
  if (v.includes("практик") || v.includes("йога")) return "🧘 Практики";
  if (v.includes("обуч") || v.includes("лекц") || v.includes("курс")) return "📚 Обучение";
  if (v.includes("игр") || v.includes("квиз")) return "🎮 Игры";
  if (v.includes("искус") || v.includes("арт")) return "🎨 Искусство";
  if (v.includes("бизнес")) return "💼 Бизнес";
  return "";
};

const normalizeMoodInput = (value: string): string => {
  const v = lower(value);
  if (!v) return "";
  if (v.includes("потус")) return "🎉 Потусить";
  if (v.includes("вдохнов")) return "✨ Вдохновиться";
  if (v.includes("отдох")) return "🧘 Отдохнуть";
  if (v.includes("поигр")) return "🎮 Поиграть";
  if (v.includes("нов")) return "🌈 Узнать новое";
  if (v.includes("природ")) return "🌿 На природу";
  if (v.includes("пообщ")) return "💛 Пообщаться";
  if (v.includes("разв")) return "📚 Развиваться";
  return "";
};

const normalizeCollectionInput = (value: string): CollectionKey | "" => {
  const v = lower(value);
  if (!v) return "";
  if (v.includes("луч") || v.includes("актуал") || v.includes("топ") || v.includes("best")) return "🔥 Лучшее и актуальное";
  if (v.includes("тело") || v.includes("движ") || v.includes("спорт") || v.includes("dance") || v.includes("йога")) return "💃 Тело и движение";
  if (v.includes("сем") || v.includes("дет")) return "👨‍👩‍👧 Семья и дети";
  if (v.includes("отнош") || v.includes("развит") || v.includes("коммуник") || v.includes("общ")) return "💛 Отношения и развитие";
  if (v.includes("отдых") || v.includes("перезагруз") || v.includes("релакс") || v.includes("природ")) return "🌿 Отдых и перезагрузка";
  if (v.includes("игр") || v.includes("досуг") || v.includes("квиз") || v.includes("настол")) return "🎲 Игры и досуг";
  return "";
};

const runtimeStatus = (e: EventItem): EventStatus => {
  if (e.status === "cancelled" || e.status === "draft") return e.status;
  const ts = parseDateTime(e);
  return ts >= Date.now() ? "upcoming" : "past";
};

const canShowUser = (e: EventItem) => runtimeStatus(e) === "upcoming" && !!e.title && !!e.date;

const seedEvents: EventItem[] = [
  { id: "e1", title: "Йога на крыше", shortDescription: "Утренняя практика йоги для перезагрузки.", description: "Практика с инструктором на крыше, подходит новичкам.", date: "2026-05-10", time: "19:00", city: "Бангкок", location: "Sky Hub", address: "Silom 17", price: "700 THB", isFree: false, category: "🧘 Практики", categories: ["🧘 Практики"], mood: "🧘 Отдохнуть", moods: ["🧘 Отдохнуть"], collections: ["🔥 Лучшее и актуальное", "🌿 Отдых и перезагрузка"], tags: ["йога"], detailsLink: "https://example.com/yoga", ticketLink: "", image: "", organizer: "Urban Flow", isVip: false, status: "upcoming", source: "seed", aiConfidence: 0.9, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }
];

function App() {
  const [events, setEvents] = useState<EventItem[]>(() => {
    const parsed = asArray<EventItem>(parse<unknown>(localStorage.getItem("event-events") ?? "[]", []), []);
    return parsed.length ? parsed : seedEvents;
  });
  const [favorites, setFavorites] = useState<string[]>(() => asArray<string>(parse<unknown>(localStorage.getItem("event-favorites") ?? "[]", []), []));
  const [collectionFavorites, setCollectionFavorites] = useState<string[]>(() =>
    asArray<string>(parse<unknown>(localStorage.getItem("event-collection-favorites") ?? "[]", []), [])
  );
  const [collectionCards, setCollectionCards] = useState<CollectionCard[]>(() => {
    const raw = asArray<unknown>(parse<unknown>(localStorage.getItem("event-collection-cards") ?? "[]", []), []);
    return raw
      .map((v) => {
        const o = asObject<Record<string, unknown>>(v, {});
        const collectionsRaw = o.collections;
        const collectionLegacy = normalize(o.collection);
        const collections =
          Array.isArray(collectionsRaw) && collectionsRaw.length
            ? (collectionsRaw as unknown[]).map((x) => normalize(x) as CollectionKey).filter((x) => (collectionKeys as readonly string[]).includes(x))
            : collectionLegacy
              ? ([collectionLegacy] as CollectionKey[]).filter((x) => (collectionKeys as readonly string[]).includes(x))
              : [];
        return {
          id: normalize(o.id),
          collections: collections.length ? collections : (["🔥 Лучшее и актуальное"] as CollectionKey[]),
          title: normalize(o.title),
          image: normalize(o.image),
          description: normalize(o.description),
          link: normalize(o.link),
          createdAt: normalize(o.createdAt) || new Date().toISOString()
        } satisfies CollectionCard;
      })
      .filter((c) => c.id && c.title);
  });
  const [report, setReport] = useState<ProcessingReport>(() => asObject<ProcessingReport>(parse<unknown>(localStorage.getItem("event-report") ?? "null", emptyReport), emptyReport));
  const [page, setPage] = useState<Page>("collections");
  const [adminSubPage] = useState<AdminSubPage>("dashboard");
  const [search, setSearch] = useState("");
  const [dateMode, setDateMode] = useState<"today" | "tomorrow" | "weekend" | "manual" | "all">("all");
  const [manualDate, setManualDate] = useState("");
  const [categoryMode, setCategoryMode] = useState("all");
  const [moodMode, setMoodMode] = useState("all");
  const [collectionMode, setCollectionMode] = useState<CollectionKey | null>(null);
  const [file, setFile] = useState<File | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [adminLogged, setAdminLogged] = useState(() => localStorage.getItem("event-admin-auth") === "1");
  const [adminLoginInput, setAdminLoginInput] = useState("");
  const [adminPasswordInput, setAdminPasswordInput] = useState("");
  const [adminError, setAdminError] = useState("");
  const [collectionCardDrafts, setCollectionCardDrafts] = useState<Record<string, CollectionCard>>({});
  const [vipDraft, setVipDraft] = useState({
    title: "",
    description: "",
    date: "",
    time: "",
    location: "",
    price: "",
    link: "",
    image: "",
    categories: [] as string[],
    moods: [] as string[]
  });
  const [vipCollections, setVipCollections] = useState<CollectionKey[]>([]);
  const [routePath, setRoutePath] = useState(() => window.location.pathname);
  const [tgGateOpen, setTgGateOpen] = useState(true);
  const [tgChecked, setTgChecked] = useState(false);
  const [tgChecking, setTgChecking] = useState(true);

  const navigate = (next: Page) => {
    setPage(next);
    setSearch("");
    if (next === "date") {
      setDateMode("all");
      setManualDate("");
    }
    if (next === "category") setCategoryMode("all");
    if (next === "mood") setMoodMode("all");
  };

  useEffect(() => { localStorage.setItem("event-events", JSON.stringify(events)); }, [events]);
  useEffect(() => { localStorage.setItem("event-favorites", JSON.stringify(favorites)); }, [favorites]);
  useEffect(() => { localStorage.setItem("event-collection-favorites", JSON.stringify(collectionFavorites)); }, [collectionFavorites]);
  useEffect(() => { localStorage.setItem("event-collection-cards", JSON.stringify(collectionCards)); }, [collectionCards]);
  useEffect(() => { localStorage.setItem("event-report", JSON.stringify(report)); }, [report]);
  useEffect(() => { localStorage.setItem("event-admin-auth", adminLogged ? "1" : "0"); }, [adminLogged]);
  useEffect(() => {
    let cancelled = false;

    const tryAutoCheck = async () => {
      setTgChecking(true);
      try {
        const stored = localStorage.getItem("event-tg-ok") === "1";

        const w = window as unknown as { Telegram?: { WebApp?: { initDataUnsafe?: { user?: { id?: unknown } } } } };
        const userIdMaybe = w.Telegram?.WebApp?.initDataUnsafe?.user?.id;
        const userId = typeof userIdMaybe === "number" ? userIdMaybe : userIdMaybe ? Number(userIdMaybe) : NaN;

        if (TELEGRAM_BOT_TOKEN && Number.isFinite(userId)) {
          const resp = await fetch(
            `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getChatMember?chat_id=${encodeURIComponent(TELEGRAM_CHAT_ID)}&user_id=${encodeURIComponent(
              String(userId)
            )}`
          );
          const data = await resp.json();

          const status = data?.result?.status;
          const subscribed = status === "member" || status === "administrator" || status === "creator";
          if (!cancelled) {
            if (subscribed) {
              setTgGateOpen(false);
              setTgChecked(true);
              localStorage.setItem("event-tg-ok", "1");
            } else {
              setTgGateOpen(true);
              setTgChecked(false);
            }
          }
          return;
        }

        // Fallback: if user confirmed earlier and we can't auto-check, allow.
        if (stored && !cancelled) {
          setTgGateOpen(false);
          setTgChecked(true);
          return;
        }
      } catch {
        // ignore - fallback UI will be shown
      } finally {
        if (!cancelled) setTgChecking(false);
      }
    };

    void tryAutoCheck();
    return () => {
      cancelled = true;
    };
  }, []);
  useEffect(() => {
    const onPop = () => setRoutePath(window.location.pathname);
    window.addEventListener("popstate", onPop);
    return () => window.removeEventListener("popstate", onPop);
  }, []);

  const isAdminRoute = routePath.startsWith("/admin");
  const navigateTo = (path: "/" | "/admin") => {
    if (routePath !== path) {
      window.history.pushState({}, "", path);
      setRoutePath(path);
    }
  };

  const allUpcomingSorted = useMemo(() => {
    return events
      .filter(canShowUser)
      .filter((e) => {
        if (!search) return true;
        const q = lower(search);
        const hay = [e.title, e.description, e.shortDescription, e.location, e.organizer, e.category, e.mood, ...e.tags];
        return hay.some((x) => lower(x).includes(q));
      })
      .sort((a, b) => parseDateTime(a) - parseDateTime(b));
  }, [events, search]);

  const byDateList = useMemo(() => {
    const now = new Date();
    const today = now.toISOString().slice(0, 10);
    const tomorrow = new Date(now.getTime() + 86400000).toISOString().slice(0, 10);
    if (dateMode === "all") return allUpcomingSorted;
    return allUpcomingSorted.filter((e) => {
      if (dateMode === "today") return e.date === today;
      if (dateMode === "tomorrow") return e.date === tomorrow;
      if (dateMode === "manual") return manualDate ? e.date === manualDate : true;
      if (dateMode === "weekend") {
        const day = new Date(`${e.date}T12:00`).getDay();
        return day === 0 || day === 6;
      }
      return true;
    });
  }, [allUpcomingSorted, dateMode, manualDate]);

  const byCategoryList = useMemo(() => {
    if (categoryMode === "all") return allUpcomingSorted;
    return allUpcomingSorted.filter((e) => e.category === categoryMode || e.categories?.includes(categoryMode));
  }, [allUpcomingSorted, categoryMode]);

  const byMoodList = useMemo(() => {
    if (moodMode === "all") return allUpcomingSorted;
    return allUpcomingSorted.filter((e) => e.mood === moodMode || e.moods?.includes(moodMode));
  }, [allUpcomingSorted, moodMode]);
  const favoriteList = useMemo(() => allUpcomingSorted.filter((e) => favorites.includes(e.id)), [allUpcomingSorted, favorites]);
  const collectionFavoriteList = useMemo(
    () => collectionCards.filter((c) => collectionFavorites.includes(c.id)),
    [collectionCards, collectionFavorites]
  );
  const vipEventsAdmin = useMemo(() => events.filter((e) => e.isVip), [events]);

  const toggleFavorite = (id: string) => setFavorites((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  const patchEvent = (id: string, patch: Partial<EventItem>) =>
    setEvents((prev) => prev.map((x) => (x.id === id ? { ...x, ...patch, updatedAt: new Date().toISOString() } : x)));
  const deleteEventById = (id: string) => setEvents((prev) => prev.filter((x) => x.id !== id));

  const readRows = async (upload: File): Promise<Record<string, unknown>[]> => {
    const ext = upload.name.split(".").pop()?.toLowerCase();
    if (ext === "json") {
      const parsed = parse<unknown>(await upload.text(), []);
      return Array.isArray(parsed) ? (parsed as Record<string, unknown>[]) : [parsed as Record<string, unknown>];
    }
    const wb = XLSX.read(await upload.arrayBuffer(), { type: "array" });
    const out: Record<string, unknown>[] = [];
    for (let si = 0; si < wb.SheetNames.length; si++) {
      const ws = wb.Sheets[wb.SheetNames[si]];
      const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" }) as unknown[][];
      const cleaned = aoa.filter((r) => Array.isArray(r) && r.some((cell) => normalize(cell)));
      if (!cleaned.length) continue;
      out.push(...aoaToSemanticRows(cleaned));
    }
    return out;
  };

  const field = (row: Record<string, unknown>, aliases: string[]) => {
    const entries = Object.entries(row);
    const normalizedAliases = aliases.map((a) => normalizeKey(a));
    for (const alias of aliases) {
      const hit = entries.find(([k]) => lower(k) === alias);
      if (hit) return normalize(hit[1]);
    }
    for (const alias of normalizedAliases) {
      const hit = entries.find(([k]) => normalizeKey(k) === alias || normalizeKey(k).includes(alias) || alias.includes(normalizeKey(k)));
      if (hit) return normalize(hit[1]);
    }
    return "";
  };
  const fieldRaw = (row: Record<string, unknown>, aliases: string[]): unknown => {
    const entries = Object.entries(row);
    const normalizedAliases = aliases.map((a) => normalizeKey(a));
    for (const alias of aliases) {
      const hit = entries.find(([k]) => lower(k) === alias);
      if (hit) return hit[1];
    }
    for (const alias of normalizedAliases) {
      const hit = entries.find(([k]) => normalizeKey(k) === alias || normalizeKey(k).includes(alias) || alias.includes(normalizeKey(k)));
      if (hit) return hit[1];
    }
    return "";
  };

  const processFile = async () => {
    if (!file) return;
    setIsLoading(true);
    const rows = await readRows(file);
    const now = new Date().toISOString();
    const rep = { ...emptyReport, loaded: rows.length };
    setEvents((prev) => {
      const next = [...prev];
      for (const row of rows) {
        const title = field(row, ["title", "название", "названиемероприятия", "eventtitle", "name"]);
        const shortFromFile = field(row, ["shortdescription", "short", "краткоеописание", "краткое описание"]);
        const extrasBlock = normalize((row as Record<string, unknown>).__extras);
        const descriptionBase = field(row, ["description", "описание"]) || shortFromFile;
        const description = extrasBlock ? (descriptionBase ? `${descriptionBase}\n${extrasBlock}` : extrasBlock) : descriptionBase;
        const parsedDate = normalizeDateInput(fieldRaw(row, ["date", "дата", "датамероприятия", "eventdate", "startdate", "датаначала", "event_start_date"]));
        const time = normalizeTimeInput(fieldRaw(row, ["time", "время", "времяначала", "eventtime", "starttime", "времястарта", "event_start_time"]));
        const date = bumpDateToFuture(parsedDate, time);
        const city = field(row, ["city", "город", "cityname"]);
        const location = field(row, ["location", "место", "venue", "площадка"]);
        const detailsLink = field(row, ["detailslink", "link", "ссылка", "url", "eventurl", "ссылканамероприятие", "event_link", "ссылкаподробности"]);
        const ticketLink = field(row, ["ticketlink", "билет", "регистрация", "ticketurl", "registrationlink", "registration_url", "ссылканарегистрацию"]);
        const sourceStatus = field(row, ["status", "статус"]);
        const tagsRaw = field(row, ["tags", "теги", "tag"]);
        const tagList = tagsRaw
          .split(/[,;\n|]+/)
          .map((x) => normalize(x))
          .filter(Boolean);

        const tagVariants = tagList.flatMap((t) => {
          const stripped = stripLeadingEmoji(t);
          return stripped && stripped !== t ? [t, stripped] : [t];
        });
        const categoriesFromTag = Array.from(new Set(tagVariants.map(normalizeCategoryInput).filter(Boolean)));
        const moodsFromTag = Array.from(new Set(tagVariants.map(normalizeMoodInput).filter(Boolean)));
        const categoryFromTag = categoriesFromTag[0] ?? "";
        const moodFromTag = moodsFromTag[0] ?? "";
        const collectionsFromTag = Array.from(
          new Set(
            tagVariants
              .map((t) => {
                if ((collectionKeys as readonly string[]).includes(t)) return t as CollectionKey;
                const inferred = normalizeCollectionInput(t);
                return inferred || "";
              })
              .filter((x): x is CollectionKey => Boolean(x))
          )
        );

        const rawText = lower(`${title} ${description} ${location}`);
        const categoryFromFile = normalizeCategoryInput(field(row, ["category", "категория"])) || categoryFromTag;
        const moodFromFile = normalizeMoodInput(field(row, ["mood", "настроение"])) || moodFromTag;
        const category = categoryFromFile || detectCategory(rawText);
        const mood = moodFromFile || detectMood(rawText);
        const categoriesResolved = Array.from(new Set([...(categoriesFromTag.length ? categoriesFromTag : []), category]));
        const moodsResolved = Array.from(new Set([...(moodsFromTag.length ? moodsFromTag : []), mood]));
        const collections = Array.from(new Set([...(collectionsFromTag.length ? collectionsFromTag : []), ...detectCollections(category, mood)]));
        const resolvedLink = detailsLink || ticketLink;
        const status = lower(sourceStatus).includes("cancel") ? "cancelled" : "upcoming";
        const exactFingerprint = `${lower(title)}|${date}|${time}|${lower(location)}|${lower(city)}|${lower(field(row, ["organizer", "организатор"]))}`;
        const dup = next.findIndex((e) => {
          const byLink = resolvedLink && (e.detailsLink === resolvedLink || e.ticketLink === resolvedLink);
          const otherFingerprint = `${lower(e.title)}|${e.date}|${e.time}|${lower(e.location)}|${lower(e.city)}|${lower(e.organizer)}`;
          const byExactFields = otherFingerprint === exactFingerprint;
          return Boolean(byLink || byExactFields);
        });
        const patch: EventItem = {
          id: crypto.randomUUID(),
          title,
          shortDescription: short(shortFromFile || description),
          description,
          date,
          time,
          city,
          location,
          address: field(row, ["address", "адрес"]),
          price: field(row, ["price", "цена"]) || "По запросу",
          isFree: lower(field(row, ["isfree", "free", "бесплатно", "price", "цена"])).includes("бесплат"),
          category,
          categories: categoriesResolved,
          mood,
          moods: moodsResolved,
          collections,
          tags: Array.from(new Set([...tagList, category, mood].map((x) => normalize(x.replace(/^.. /, ""))).filter(Boolean))),
          detailsLink: detailsLink || ticketLink,
          ticketLink,
          image: field(row, ["image", "изображение"]),
          organizer: field(row, ["organizer", "организатор"]),
          isVip: false,
          status,
          source: file.name,
          aiConfidence: title && date ? 0.9 : 0.65,
          createdAt: now,
          updatedAt: now
        };
        if (dup >= 0) { rep.duplicates += 1; rep.updated += 1; next[dup] = { ...next[dup], ...patch, id: next[dup].id, createdAt: next[dup].createdAt }; } else { rep.added += 1; next.push(patch); }
      }
      rep.hiddenPast = next.filter((e) => runtimeStatus(e) === "past").length;
      rep.cancelled = next.filter((e) => e.status === "cancelled").length;
      rep.review = next.filter((e) => e.status === "needs_review").length;
      rep.published = next.filter(canShowUser).length;
      return next;
    });
    setReport(rep);
    setIsLoading(false);
  };

  const onCollectionImageUpload = (fileInput: File | null, onDone: (imageDataUrl: string) => void) => {
    if (!fileInput) return;
    const reader = new FileReader();
    reader.onload = () => {
      const src = typeof reader.result === "string" ? reader.result : "";
      if (!src) return;

      const img = new Image();
      img.onload = () => {
        const targetW = 1600;
        const targetH = 900; // 16:9
        const scale = Math.min(targetW / img.width, targetH / img.height);
        const w = Math.max(1, Math.round(img.width * scale));
        const h = Math.max(1, Math.round(img.height * scale));
        const x = Math.round((targetW - w) / 2);
        const y = Math.round((targetH - h) / 2);

        const canvas = document.createElement("canvas");
        canvas.width = targetW;
        canvas.height = targetH;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          onDone(src);
          return;
        }

        ctx.fillStyle = "#fff9e6";
        ctx.fillRect(0, 0, targetW, targetH);
        ctx.drawImage(img, x, y, w, h);

        const out = canvas.toDataURL("image/jpeg", 0.85);
        onDone(out);
      };
      img.onerror = () => onDone(src);
      img.src = src;
    };
    reader.readAsDataURL(fileInput);
  };

  const clearUploadedData = () => {
    setEvents((prev) => prev.filter((e) => e.source === "seed"));
    setCollectionCards([]);
    setReport(emptyReport);
    setFile(null);
  };

  const updateCollectionCard = (id: string, patch: Partial<CollectionCard>) => {
    setCollectionCards((prev) =>
      prev.map((c) =>
        c.id === id
          ? {
              ...c,
              ...patch,
              description: patch.description !== undefined ? patch.description.slice(0, 250) : c.description
            }
          : c
      )
    );
  };
  const deleteCollectionCard = (id: string) => {
    setCollectionCards((prev) => prev.filter((c) => c.id !== id));
    setCollectionFavorites((prev) => prev.filter((x) => x !== id));
    setCollectionCardDrafts((prev) => {
      if (!prev[id]) return prev;
      const { [id]: _removed, ...rest } = prev;
      return rest;
    });
  };

  const getCollectionCardDraft = (card: CollectionCard) => collectionCardDrafts[card.id] ?? card;
  const setCollectionCardDraft = (id: string, patch: Partial<CollectionCard>) => {
    setCollectionCardDrafts((prev) => {
      const base = prev[id] ?? collectionCards.find((x) => x.id === id);
      if (!base) return prev;
      return {
        ...prev,
        [id]: {
          ...base,
          ...patch,
          description: patch.description !== undefined ? patch.description.slice(0, 250) : base.description
        }
      };
    });
  };
  const discardCollectionCardDraft = (id: string) => {
    setCollectionCardDrafts((prev) => {
      if (!prev[id]) return prev;
      const { [id]: _removed, ...rest } = prev;
      return rest;
    });
  };
  const saveCollectionCardDraft = (id: string) => {
    const draft = collectionCardDrafts[id];
    if (!draft) return;
    updateCollectionCard(id, {
      collections: draft.collections,
      title: draft.title,
      description: draft.description,
      link: draft.link,
      image: draft.image
    });
    discardCollectionCardDraft(id);
  };

  const cardsByCollection =
    collectionMode === null ? collectionCards : collectionCards.filter((c) => c.collections.includes(collectionMode));

  const vipEventsByCollection = useMemo(() => {
    const vip = allUpcomingSorted.filter((e) => e.isVip);
    if (collectionMode === null) return vip;
    return vip.filter((e) => e.collections.includes(collectionMode));
  }, [allUpcomingSorted, collectionMode]);

  const dateParts = (iso: string) => {
    const dt = new Date(`${iso}T12:00`);
    if (Number.isNaN(dt.getTime())) return { day: "--", month: "---" };
    return {
      day: String(dt.getDate()),
      month: dt.toLocaleDateString("ru-RU", { month: "short" }).replace(".", "")
    };
  };

  const activeTabClass = (tab: Page) => (page === tab ? "dark" : "");

  const renderCards = (list: EventItem[]) => (
    <section className="list">
      <div className="scroll-nav">
        <button type="button" className="ghost" onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}>
          ↑ Вверх
        </button>
        <button type="button" className="ghost" onClick={() => window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" })}>
          ↓ Вниз
        </button>
      </div>
      {list.length === 0 && <div className="panel">Пока нет событий по этому запросу.</div>}
      {list.map((e) => {
        const wd = weekdayShortRu(e.date);
        return (
        <article key={e.id} className="card">
          <div className="date-badge">
            <b>{dateParts(e.date).day}</b>
            <span>{dateParts(e.date).month}</span>
            <small>{e.time || "--:--"}</small>
            {wd && <span className="weekday-note">({wd})</span>}
          </div>
          <div className="card-side">
            {e.detailsLink || e.ticketLink ? (
              <a href={e.detailsLink || e.ticketLink} target="_blank" rel="noreferrer">
                Подробнее
              </a>
            ) : (
              <button className="disabled-btn" disabled>
                Подробнее
              </button>
            )}
            <button className="ghost" onClick={() => toggleFavorite(e.id)}>
              {favorites.includes(e.id) ? "Убрать" : "В избранное"}
            </button>
          </div>
          <div className="card-main">
            {e.isVip && <span className="vip-badge">VIP</span>}
            {e.image && <img className="event-cover" src={e.image} alt={e.title} />}
            <h3>{e.title}</h3>
            <p className="clamp">{e.shortDescription}</p>
            <small>
              {[normalize(e.city), normalize(e.location)].filter(Boolean).join(" • ") || "Место уточняется"} • {e.price || "По запросу"}
            </small>
          </div>
        </article>
        );
      })}
      <div className="scroll-nav bottom">
        <button type="button" className="ghost" onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}>
          ↑ Вверх
        </button>
        <button type="button" className="ghost" onClick={() => window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" })}>
          ↓ Вниз
        </button>
      </div>
    </section>
  );

  const adminContent = (
    <div className="page">
      <header className="panel hero">
        <h1>Админ-панель</h1>
        <p>Управление базой мероприятий, импортом и подборками</p>
        <div className="actions menu-actions">
          <button onClick={() => navigateTo("/")}>Перейти на сайт</button>
          {adminLogged && <button className="ghost" onClick={() => setAdminLogged(false)}>Выйти</button>}
        </div>
      </header>

      {!adminLogged ? (
        <section className="panel">
          <h2>Вход по логину и паролю</h2>
          <input placeholder="Логин" value={adminLoginInput} onChange={(e) => setAdminLoginInput(e.target.value)} />
          <input placeholder="Пароль" type="password" value={adminPasswordInput} onChange={(e) => setAdminPasswordInput(e.target.value)} />
          {adminError && <p className="error-text">{adminError}</p>}
          <button
            onClick={() => {
              if (adminLoginInput === ADMIN_LOGIN && adminPasswordInput === ADMIN_PASSWORD) {
                setAdminLogged(true);
                setAdminError("");
                setAdminPasswordInput("");
              } else {
                setAdminError("Неверный логин или пароль");
              }
            }}
          >
            Войти
          </button>
        </section>
      ) : (
        <>
          {adminSubPage === "dashboard" && (
            <>
              <section className="panel">
                <h2>Загрузка CSV/XLSX/JSON</h2>
                <input type="file" accept=".csv,.xlsx,.json" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
                <div className="actions">
                  <button onClick={() => void processFile()} disabled={!file || isLoading}>{isLoading ? "Обработка..." : "Запустить ИИ-обработку"}</button>
                  <button className="ghost" onClick={clearUploadedData}>Удалить загруженные данные</button>
                </div>
                <p>Загружено: {report.loaded} | Добавлено: {report.added} | Обновлено: {report.updated} | Дубли: {report.duplicates} | Опубликовано: {report.published}</p>
              </section>
              <section className="panel">
                <h3>VIP‑мероприятие в топ‑подборках (вручную)</h3>
                <p className="hint">Заполните поля: дата, время, место, цена, заголовок, описание, ссылка и фото — событие появится в выбранных подборках с пометкой VIP.</p>
                <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))" }}>
                  <input type="date" value={vipDraft.date} onChange={(e) => setVipDraft((p) => ({ ...p, date: e.target.value }))} />
                  <input placeholder="Время (например 19:00)" value={vipDraft.time} onChange={(e) => setVipDraft((p) => ({ ...p, time: e.target.value }))} />
                  <input placeholder="Место / площадка" value={vipDraft.location} onChange={(e) => setVipDraft((p) => ({ ...p, location: e.target.value }))} />
                  <input placeholder="Цена" value={vipDraft.price} onChange={(e) => setVipDraft((p) => ({ ...p, price: e.target.value }))} />
                  <input placeholder="Заголовок" value={vipDraft.title} onChange={(e) => setVipDraft((p) => ({ ...p, title: e.target.value }))} />
                  <input placeholder="Ссылка (подробнее / билет)" value={vipDraft.link} onChange={(e) => setVipDraft((p) => ({ ...p, link: e.target.value }))} />
                </div>
                <textarea placeholder="Описание" value={vipDraft.description} onChange={(e) => setVipDraft((p) => ({ ...p, description: e.target.value }))} />
                <div className="actions">
                  <label className="ghost" style={{ display: "inline-flex", alignItems: "center", gap: 10 }}>
                    Загрузить фото (16:9)
                    <input
                      type="file"
                      accept="image/*"
                      style={{ display: "none" }}
                      onChange={(e) => onCollectionImageUpload(e.target.files?.[0] ?? null, (image) => setVipDraft((p) => ({ ...p, image })))}
                    />
                  </label>
                  {vipDraft.image && <button className="ghost" onClick={() => setVipDraft((p) => ({ ...p, image: "" }))}>Убрать фото</button>}
                </div>
                {vipDraft.image && <img className="cover preview-cover" src={vipDraft.image} alt="VIP фото" />}

                <h3 style={{ marginTop: 10 }}>Теги (категории)</h3>
                <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))" }}>
                  {categories.map((c) => {
                    const checked = vipDraft.categories.includes(c);
                    return (
                      <label key={c} className="ghost" style={{ display: "flex", gap: 10, alignItems: "center", padding: 10, borderRadius: 10 }}>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(e) =>
                            setVipDraft((p) => ({
                              ...p,
                              categories: e.target.checked ? Array.from(new Set([...p.categories, c])) : p.categories.filter((x) => x !== c)
                            }))
                          }
                        />
                        <span>{c}</span>
                      </label>
                    );
                  })}
                </div>

                <h3 style={{ marginTop: 10 }}>Теги (настроение)</h3>
                <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))" }}>
                  {moods.map((m) => {
                    const checked = vipDraft.moods.includes(m);
                    return (
                      <label key={m} className="ghost" style={{ display: "flex", gap: 10, alignItems: "center", padding: 10, borderRadius: 10 }}>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(e) =>
                            setVipDraft((p) => ({
                              ...p,
                              moods: e.target.checked ? Array.from(new Set([...p.moods, m])) : p.moods.filter((x) => x !== m)
                            }))
                          }
                        />
                        <span>{m}</span>
                      </label>
                    );
                  })}
                </div>

                <h3 style={{ marginTop: 10 }}>Разместить в подборках</h3>
                <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))" }}>
                  {collectionKeys.map((k) => {
                    const checked = vipCollections.includes(k);
                    return (
                      <label key={k} className="ghost" style={{ display: "flex", gap: 10, alignItems: "center", padding: 10, borderRadius: 10 }}>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(e) =>
                            setVipCollections((prev) => (e.target.checked ? Array.from(new Set([...prev, k])) : prev.filter((x) => x !== k)))
                          }
                        />
                        <span>{k}</span>
                      </label>
                    );
                  })}
                </div>

                <div className="actions">
                  <button
                    onClick={() => {
                      const now = new Date().toISOString();
                      const title = normalize(vipDraft.title);
                      const date = normalize(vipDraft.date);
                      if (!title || !date) return;
                      const categoriesPicked = Array.from(new Set(vipDraft.categories));
                      const moodsPicked = Array.from(new Set(vipDraft.moods));
                      const primaryCategory = categoriesPicked[0] || detectCategory(lower(`${title} ${vipDraft.description} ${vipDraft.location}`));
                      const primaryMood = moodsPicked[0] || detectMood(lower(`${title} ${vipDraft.description} ${vipDraft.location}`));
                      const collectionsPicked = Array.from(new Set(vipCollections.length ? vipCollections : ["🔥 Лучшее и актуальное"]));
                      const catsResolved = Array.from(new Set([primaryCategory, ...categoriesPicked]));
                      const moodsResolved = Array.from(new Set([primaryMood, ...moodsPicked]));
                      setEvents((prev) => [
                        {
                          id: crypto.randomUUID(),
                          title,
                          shortDescription: short(vipDraft.description),
                          description: normalize(vipDraft.description),
                          date: bumpDateToFuture(date, normalize(vipDraft.time)),
                          time: normalizeTimeInput(vipDraft.time),
                          city: "",
                          location: normalize(vipDraft.location),
                          address: "",
                          price: normalize(vipDraft.price) || "По запросу",
                          isFree: lower(vipDraft.price).includes("бесплат"),
                          category: primaryCategory,
                          categories: catsResolved,
                          mood: primaryMood,
                          moods: moodsResolved,
                          collections: collectionsPicked,
                          tags: tagsFromCategoriesAndMoods(catsResolved, moodsResolved),
                          detailsLink: normalize(vipDraft.link),
                          ticketLink: "",
                          image: normalize(vipDraft.image),
                          organizer: "VIP",
                          isVip: true,
                          status: "upcoming",
                          source: "vip",
                          aiConfidence: 0.99,
                          createdAt: now,
                          updatedAt: now
                        },
                        ...prev
                      ]);
                      setVipDraft({
                        title: "",
                        description: "",
                        date: "",
                        time: "",
                        location: "",
                        price: "",
                        link: "",
                        image: "",
                        categories: [],
                        moods: []
                      });
                      setVipCollections([]);
                    }}
                    disabled={!normalize(vipDraft.title) || !normalize(vipDraft.date)}
                  >
                    Добавить мероприятие
                  </button>
                </div>
              </section>
              <section className="panel">
                <h3>Топ‑подборки: редактирование и удаление</h3>
                <p className="hint" style={{ marginBottom: 12 }}>
                  <strong>VIP‑мероприятия</strong> (форма выше) сохраняются как события и показываются здесь. Блок ниже — только отдельные промо‑карточки без даты.
                </p>

                <h4 style={{ margin: "0 0 8px", fontSize: 16 }}>VIP‑мероприятия</h4>
                {vipEventsAdmin.length === 0 ? (
                  <p style={{ marginBottom: 18 }}>Пока нет VIP‑мероприятий — добавьте через форму «VIP‑мероприятие в топ‑подборках».</p>
                ) : (
                  <div className="list" style={{ marginTop: 0, marginBottom: 22 }}>
                    {vipEventsAdmin.map((e) => (
                      <details key={e.id} className="panel" style={{ padding: 14 }}>
                        <summary style={{ cursor: "pointer" }}>
                          <b>{e.title || "Без названия"}</b>{" "}
                          <span style={{ opacity: 0.7 }}>
                            — {e.date} {e.time || ""} • VIP
                          </span>
                        </summary>
                        <div style={{ marginTop: 12, display: "grid", gap: 10 }}>
                          <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))" }}>
                            <input placeholder="Заголовок" value={e.title} onChange={(ev) => patchEvent(e.id, { title: ev.target.value })} />
                            <input type="date" value={e.date} onChange={(ev) => patchEvent(e.id, { date: ev.target.value })} />
                            <input placeholder="Время" value={e.time} onChange={(ev) => patchEvent(e.id, { time: normalizeTimeInput(ev.target.value) })} />
                            <input placeholder="Место / площадка" value={e.location} onChange={(ev) => patchEvent(e.id, { location: ev.target.value })} />
                            <input placeholder="Цена" value={e.price} onChange={(ev) => patchEvent(e.id, { price: ev.target.value })} />
                            <input placeholder="Ссылка" value={e.detailsLink} onChange={(ev) => patchEvent(e.id, { detailsLink: ev.target.value })} />
                          </div>
                          <textarea
                            placeholder="Описание"
                            value={e.description}
                            onChange={(ev) => {
                              const v = ev.target.value;
                              patchEvent(e.id, { description: v, shortDescription: short(v) });
                            }}
                          />
                          <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))" }}>
                            {categories.map((c) => {
                              const checked = e.categories.includes(c);
                              return (
                                <label key={`${e.id}-cat-${c}`} className="ghost" style={{ display: "flex", gap: 10, alignItems: "center", padding: 8, borderRadius: 10 }}>
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={(ev) => {
                                      let nextCats = ev.target.checked ? Array.from(new Set([...e.categories, c])) : e.categories.filter((x) => x !== c);
                                      if (!nextCats.length) nextCats = [e.category];
                                      patchEvent(e.id, {
                                        categories: nextCats,
                                        category: nextCats[0],
                                        tags: tagsFromCategoriesAndMoods(nextCats, e.moods)
                                      });
                                    }}
                                  />
                                  <span>{c}</span>
                                </label>
                              );
                            })}
                          </div>
                          <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))" }}>
                            {moods.map((m) => {
                              const checked = e.moods.includes(m);
                              return (
                                <label key={`${e.id}-mood-${m}`} className="ghost" style={{ display: "flex", gap: 10, alignItems: "center", padding: 8, borderRadius: 10 }}>
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={(ev) => {
                                      let nextMoods = ev.target.checked ? Array.from(new Set([...e.moods, m])) : e.moods.filter((x) => x !== m);
                                      if (!nextMoods.length) nextMoods = [e.mood];
                                      patchEvent(e.id, {
                                        moods: nextMoods,
                                        mood: nextMoods[0],
                                        tags: tagsFromCategoriesAndMoods(e.categories, nextMoods)
                                      });
                                    }}
                                  />
                                  <span>{m}</span>
                                </label>
                              );
                            })}
                          </div>
                          <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))" }}>
                            {collectionKeys.map((k) => {
                              const checked = e.collections.includes(k);
                              return (
                                <label key={`${e.id}-col-${k}`} className="ghost" style={{ display: "flex", gap: 10, alignItems: "center", padding: 8, borderRadius: 10 }}>
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={(ev) => {
                                      const nextCols = ev.target.checked ? Array.from(new Set([...e.collections, k])) : e.collections.filter((x) => x !== k);
                                      patchEvent(e.id, {
                                        collections: nextCols.length ? nextCols : ["🔥 Лучшее и актуальное"]
                                      });
                                    }}
                                  />
                                  <span>{k}</span>
                                </label>
                              );
                            })}
                          </div>
                          <div className="actions" style={{ flexWrap: "wrap" }}>
                            <label className="ghost" style={{ display: "inline-flex", alignItems: "center", gap: 10 }}>
                              Заменить фото
                              <input
                                type="file"
                                accept="image/*"
                                style={{ display: "none" }}
                                onChange={(ev) => onCollectionImageUpload(ev.target.files?.[0] ?? null, (image) => patchEvent(e.id, { image }))}
                              />
                            </label>
                            {e.image ? (
                              <button type="button" className="ghost" onClick={() => patchEvent(e.id, { image: "" })}>
                                Убрать фото
                              </button>
                            ) : null}
                            <button type="button" onClick={() => deleteEventById(e.id)}>
                              Удалить мероприятие
                            </button>
                          </div>
                          {e.image ? <img className="cover preview-cover" src={e.image} alt="" /> : null}
                        </div>
                      </details>
                    ))}
                  </div>
                )}

                <h4 style={{ margin: "0 0 8px", fontSize: 16 }}>Промо‑карточки подборок</h4>
                <p className="hint" style={{ marginBottom: 10 }}>
                  Отдельные карточки с картинкой и текстом без привязки к дате. Если вы добавляете только VIP‑события, здесь может не быть записей.
                </p>
                {collectionCards.length === 0 ? (
                  <p>Промо‑карточек нет.</p>
                ) : (
                  <div className="list" style={{ marginTop: 0 }}>
                    {collectionCards.map((c) => (
                      <details key={c.id} className="panel" style={{ padding: 14 }}>
                        <summary style={{ cursor: "pointer" }}>
                          <b>{c.title || "Без названия"}</b>{" "}
                          <span style={{ opacity: 0.7 }}>— {c.collections.join(" • ")}</span>
                        </summary>
                        <div style={{ marginTop: 12 }} className="card collection-card">
                          {c.image ? <img className="cover" src={c.image} alt={c.title} /> : <div className="cover cover-placeholder">16:9</div>}
                          <div className="collection-content">
                            {(() => {
                              const d = getCollectionCardDraft(c);
                              const isDirty = Boolean(collectionCardDrafts[c.id]);
                              return (
                                <>
                            <div className="actions" style={{ display: "grid", gap: 8, gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))" }}>
                              {collectionKeys.map((k) => {
                                const checked = d.collections.includes(k);
                                return (
                                  <label key={k} className="ghost" style={{ display: "flex", gap: 10, alignItems: "center", padding: 10, borderRadius: 10 }}>
                                    <input
                                      type="checkbox"
                                      checked={checked}
                                      onChange={(e) =>
                                        setCollectionCardDraft(c.id, {
                                          collections: e.target.checked ? Array.from(new Set([...d.collections, k])) : d.collections.filter((x) => x !== k)
                                        })
                                      }
                                    />
                                    <span>{k}</span>
                                  </label>
                                );
                              })}
                            </div>
                            <input
                              placeholder="Заголовок"
                              value={d.title}
                              onChange={(e) => setCollectionCardDraft(c.id, { title: e.target.value })}
                            />
                            <textarea
                              placeholder="Описание (до 250 символов)"
                              value={d.description}
                              maxLength={250}
                              onChange={(e) => setCollectionCardDraft(c.id, { description: e.target.value })}
                            />
                            <input
                              placeholder="Ссылка подробнее"
                              value={d.link}
                              onChange={(e) => setCollectionCardDraft(c.id, { link: e.target.value })}
                            />
                            <div className="actions" style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                              <label className="ghost" style={{ display: "inline-flex", alignItems: "center", gap: 10 }}>
                                Заменить картинку
                                <input
                                  type="file"
                                  accept="image/*"
                                  style={{ display: "none" }}
                                  onChange={(e) =>
                                    onCollectionImageUpload(e.target.files?.[0] ?? null, (image) => setCollectionCardDraft(c.id, { image }))
                                  }
                                />
                              </label>
                              <button className="ghost" onClick={() => setCollectionCardDraft(c.id, { image: "" })} disabled={!d.image}>
                                Удалить картинку
                              </button>
                              <button className={isDirty ? "" : "ghost"} onClick={() => saveCollectionCardDraft(c.id)} disabled={!isDirty}>
                                Сохранить
                              </button>
                              <button className="ghost" onClick={() => discardCollectionCardDraft(c.id)} disabled={!isDirty}>
                                Отменить
                              </button>
                              <button onClick={() => deleteCollectionCard(c.id)}>Удалить карточку</button>
                            </div>
                                </>
                              );
                            })()}
                          </div>
                        </div>
                      </details>
                    ))}
                  </div>
                )}
              </section>
            </>
          )}
        </>
      )}
    </div>
  );

  if (isAdminRoute) return adminContent;

  if (tgGateOpen) {
    return (
      <div className="page gate-page">
        <section className="gate-card" aria-labelledby="gate-title">
          <div className="gate-accent" aria-hidden />
          <p className="gate-eyebrow">Telegram</p>
          <h2 id="gate-title" className="gate-title">
            Доступ к подборке событий
          </h2>
          <p className="gate-lead">
            {tgChecking
              ? "Проверяем, что вы подписаны на наш канал — это займёт несколько секунд."
              : "Чтобы открыть афишу, подпишитесь на канал. Это помогает держать в курсе лучших мероприятий."}
          </p>

          {tgChecking ? (
            <div className="gate-loading" role="status" aria-live="polite">
              <span className="gate-spinner" aria-hidden />
              <span className="gate-loading-text">Проверка подписки…</span>
            </div>
          ) : null}

          <div className={`gate-body ${tgChecking ? "gate-body-wait" : ""}`.trim()}>
            <a className="gate-channel" href={TELEGRAM_CHANNEL_URL} target="_blank" rel="noreferrer">
              Перейти в канал
            </a>

            <div className="gate-confirm">
              <label className="gate-check-label">
                <input
                  type="checkbox"
                  className="gate-checkbox"
                  checked={tgChecked}
                  disabled={tgChecking}
                  onChange={(e) => setTgChecked(e.target.checked)}
                />
                <span className="gate-check-text">Я подписан(а) на канал</span>
              </label>
            </div>

            <button
              type="button"
              className="gate-continue"
              disabled={!tgChecked || tgChecking}
              onClick={() => {
                localStorage.setItem("event-tg-ok", "1");
                setTgGateOpen(false);
              }}
            >
              Войти в афишу
            </button>
          </div>

          <p className="gate-footnote">Канал откроется в новой вкладке. Затем вернитесь сюда и подтвердите вход.</p>
        </section>
      </div>
    );
  }

  return (
    <div className="page">
      <header className="panel hero">
        <h1>Как будем искать события?</h1>
        <p>Выберите формат поиска и получите актуальные мероприятия</p>
        <div className="actions menu-actions">
          <div className="menu-row menu-row-home">
            <button type="button" className={`menu-home ${activeTabClass("collections")}`.trim()} onClick={() => navigate("collections")}>
              Топ подборки (главная)
            </button>
          </div>
          <div className="menu-row menu-row-3">
            <button type="button" className={activeTabClass("date")} onClick={() => navigate("date")}>По дате</button>
            <button type="button" className={activeTabClass("category")} onClick={() => navigate("category")}>По категориям</button>
            <button type="button" className={activeTabClass("mood")} onClick={() => navigate("mood")}>По настроению</button>
          </div>
          <div className="menu-row menu-row-3">
            <button type="button" className={activeTabClass("favorites")} onClick={() => navigate("favorites")}>
              Избранное ({favorites.length + collectionFavorites.length})
            </button>
            <button type="button" className={activeTabClass("about")} onClick={() => navigate("about")}>О проекте</button>
            <button type="button" className="dark" onClick={() => navigateTo("/admin")}>Админка</button>
          </div>
        </div>
      </header>

      <section className="panel">
        <div className="search-row">
          <div className="search-box">
            <input
              placeholder="Поиск по названию, месту, организатору..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              onKeyDown={(e) => {
              if (e.key === "Enter") setPage("collections");
              }}
            />
            {search && (
              <button className="search-clear" onClick={() => setSearch("")} aria-label="Очистить поиск">
                ×
              </button>
            )}
          </div>
          <button onClick={() => setPage("collections")}>Найти</button>
        </div>
      </section>

      {page === "all" && renderCards(allUpcomingSorted)}
      {page === "date" && <><section className="panel actions"><button className={dateMode === "all" ? "dark" : "ghost"} onClick={() => setDateMode("all")}>Все актуальные</button><button className={dateMode === "today" ? "dark" : "ghost"} onClick={() => setDateMode("today")}>Сегодня</button><button className={dateMode === "tomorrow" ? "dark" : "ghost"} onClick={() => setDateMode("tomorrow")}>Завтра</button><button className={dateMode === "weekend" ? "dark" : "ghost"} onClick={() => setDateMode("weekend")}>На выходных</button><button className={dateMode === "manual" ? "dark" : "ghost"} onClick={() => setDateMode("manual")}>Выбрать дату</button></section>{dateMode === "manual" && <section className="panel"><input type="date" value={manualDate} onChange={(e) => setManualDate(e.target.value)} /></section>}{renderCards(byDateList)}</>}
      {page === "category" && <><section className="panel actions"><button className={categoryMode === "all" ? "dark" : "ghost"} onClick={() => setCategoryMode("all")}>Все категории</button>{categories.map((c) => <button key={c} className={categoryMode === c ? "dark" : "ghost"} onClick={() => setCategoryMode(c)}>{c}</button>)}</section>{renderCards(byCategoryList)}</>}
      {page === "mood" && <><section className="panel actions"><button className={moodMode === "all" ? "dark" : "ghost"} onClick={() => setMoodMode("all")}>Все настроения</button>{moods.map((m) => <button key={m} className={moodMode === m ? "dark" : "ghost"} onClick={() => setMoodMode(m)}>{m}</button>)}</section>{renderCards(byMoodList)}</>}
      {page === "collections" && (
        <>
          <section className="panel actions">
            <button className={collectionMode === null ? "dark" : "ghost"} onClick={() => setCollectionMode(null)}>
              Все подборки
            </button>
            {collectionKeys.map((c) => (
              <button key={c} className={collectionMode === c ? "dark" : "ghost"} onClick={() => setCollectionMode(c)}>
                {c}
              </button>
            ))}
          </section>
          {vipEventsByCollection.length > 0 && (
            <section className="panel">
              <h3>VIP‑мероприятия</h3>
              {renderCards(vipEventsByCollection)}
            </section>
          )}
          <section className="list">
            {cardsByCollection.length === 0 && <div className="panel">В этой подборке пока нет добавленных карточек.</div>}
            {cardsByCollection.map((c) => {
              const fav = collectionFavorites.includes(c.id);
              return (
                <article key={c.id} className="card collection-card">
                  {c.image ? <img className="cover" src={c.image} alt={c.title} /> : <div className="cover cover-placeholder">16:9</div>}
                  <div className="collection-content">
                    <h3>{c.title}</h3>
                    <p>
                      {c.description.length > 250 ? `${c.description.slice(0, 250)}...` : c.description}
                    </p>
                    <div className="collection-actions">
                      {c.link ? (
                        <a href={c.link} target="_blank" rel="noreferrer">
                          Подробнее
                        </a>
                      ) : (
                        <button className="disabled-btn" disabled>
                          Подробнее
                        </button>
                      )}
                      <button
                        className={fav ? "dark" : "ghost"}
                        onClick={() => setCollectionFavorites((prev) => (prev.includes(c.id) ? prev.filter((x) => x !== c.id) : [...prev, c.id]))}
                      >
                        {fav ? "Убрать" : "В избранное"}
                      </button>
                    </div>
                  </div>
                </article>
              );
            })}
          </section>
        </>
      )}
      {page === "favorites" && (
        <>
          {favoriteList.length === 0 && collectionFavoriteList.length === 0 && (
            <section className="panel">
              <p>Пока ничего не сохранено. Добавляйте интересные события, чтобы не потерять их.</p>
            </section>
          )}
          {favoriteList.length > 0 && renderCards(favoriteList)}
          {collectionFavoriteList.length > 0 && (
            <section className="panel">
              <h3>Топ-подборки в избранном</h3>
              <div className="list" style={{ marginTop: 0 }}>
                {collectionFavoriteList.map((c) => {
                  const fav = collectionFavorites.includes(c.id);
                  return (
                    <article key={c.id} className="card collection-card">
                      {c.image ? <img className="cover" src={c.image} alt={c.title} /> : <div className="cover cover-placeholder">16:9</div>}
                      <div className="collection-content">
                        <h3>{c.title}</h3>
                    <p>
                      {c.description.length > 250 ? `${c.description.slice(0, 250)}...` : c.description}
                    </p>
                        <div className="collection-actions">
                          {c.link ? (
                            <a href={c.link} target="_blank" rel="noreferrer">
                              Подробнее
                            </a>
                          ) : (
                            <button className="disabled-btn" disabled>
                              Подробнее
                            </button>
                          )}
                          <button
                            className={fav ? "dark" : "ghost"}
                            onClick={() => setCollectionFavorites((prev) => (prev.includes(c.id) ? prev.filter((x) => x !== c.id) : [...prev, c.id]))}
                          >
                            {fav ? "Убрать" : "В избранное"}
                          </button>
                        </div>
                      </div>
                    </article>
                  );
                })}
              </div>
            </section>
          )}
        </>
      )}
      {page === "about" && (
        <section className="panel about-page">
          <h2>О проекте</h2>
          <p className="about-lead">
            Мы собираем живые события города в одном месте: чтобы проще выбирать, куда пойти сегодня и на выходных.
          </p>

          <h3>О нас</h3>
          <p>
            Команда афиши формирует подборки мероприятий — от спокойного отдыха до активностей и тусовок. Мы следим за актуальностью
            дат и не показываем прошедшие события в основной выдаче.
          </p>

          <h3>Миссия</h3>
          <p>
            Сделать так, чтобы люди чаще выходили из дома навстречу впечатлениям: помогать находить своё настроение по дате, тематике и
            интересам — без лишнего шума и устаревших анонсов.
          </p>

          <h3>Связь с организаторами</h3>
          <p>
            Вопросы по размещению, правкам анонса и сотрудничеству пишите в Telegram:{" "}
            <a href={TELEGRAM_CONTACT_URL} target="_blank" rel="noreferrer">
              {TELEGRAM_CONTACT_HANDLE}
            </a>
            . Мы ответим в рабочем порядке и поможем корректно отразить ваше мероприятие на сайте.
          </p>

          <h3>Для организаторов</h3>
          <p>
            <strong>Добавление вашего мероприятия в наш проект — бесплатно.</strong> Расскажите о дате, месте, цене и дайте ссылку на
            подробности — мы поможем попасть в нужные подборки и к нужной аудитории.
          </p>
        </section>
      )}

    </div>
  );
}

export default App;
