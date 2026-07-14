import {
  ArrowLeft,
  BookmarkCheck,
  CalendarClock,
  Check,
  ChevronDown,
  ChevronUp,
  BadgeCheck,
  Ban,
  CircleCheck,
  CircleEllipsis,
  CircleHelp,
  CircleX,
  Bot,
  BookOpen,
  Briefcase,
  BrainCircuit,
  Download,
  FileText,
  FileInput,
  Flag,
  GraduationCap,
  Hourglass,
  Info,
  Import,
  KeyRound,
  Loader2,
  Mic,
  MoveRight,
  Newspaper,
  PartyPopper,
  Pencil,
  Plane,
  Plus,
  RefreshCcw,
  Repeat2,
  Settings,
  Sparkles,
  Stethoscope,
  Trash2,
  Utensils,
  UserRound,
  Volume2,
  X,
  Zap,
} from "lucide-react";
import { useEffect, useMemo, useRef, useState, type FormEvent, type ReactNode } from "react";
import {
  builtinDecks,
  getDeckById,
  type BuiltinDeckId,
  type VocabularyWord,
} from "./data/builtinDecks";
import * as AlertDialog from "@radix-ui/react-alert-dialog";
import * as Dialog from "@radix-ui/react-dialog";
import * as Popover from "@radix-ui/react-popover";
import * as Select from "@radix-ui/react-select";
import * as Slider from "@radix-ui/react-slider";
import * as XLSX from "xlsx";
import {
  getSupabaseSession,
  signInAnonymously,
  signInWithPassword,
  signOutSupabase,
  signUpWithPassword,
  supabase,
} from "./lib/supabase";
import { loadSupabaseData, syncLocalDataToSupabase, type SupabaseHydrationData } from "./lib/supabaseSync";

type ButtonTone = "primary" | "outline" | "soft" | "neutral";
type ButtonSize = "sm" | "md";
type ReviewState = "notStarted" | "inProgress" | "todayDone" | "allDone";
type View = "home" | "study" | "translate" | "translateResult" | "review";
type MarkStatus = "known" | "familiar" | "unknown";
type ReviewDeckFilter = string | "all";
type TranslationMode = "enToZh" | "zhToEn";
type TranslationStatus = "idle" | "generating" | "evaluating";
type ImportDeckTab = "ai" | "file";
type DeckConfirmAction = "reset" | "delete";
type AuthMode = "signIn" | "signUp";

type AppSettings = {
  deepseekApiKey: string;
  dailyStudyCount: number;
  dailyReviewGoal: number;
};

type ReviewData = {
  state: ReviewState;
  status: string;
  percent: number;
  due: number;
  overdue: number;
  completed: number;
};

type CalendarDay = {
  label: string;
  status: "muted" | "done" | "missed" | "today" | "upcoming";
};

type StudySession = {
  id: string;
  deckId: string;
  wordIds: string[];
  statuses: Record<string, MarkStatus>;
  createdAt: string;
};

type VocabularyDeck = {
  id: string;
  name: string;
  description: string;
  words: VocabularyWord[];
  aiGenerated?: boolean;
  custom?: boolean;
};

type StoredProgress = Record<string, Record<string, MarkStatus>>;
type StoredMastery = Record<string, Record<string, true>>;
type StoredReviewCycles = Record<string, Record<string, number>>;
type StoredReviewStats = Record<string, number>;
type StoredDailyReviewedWords = Record<string, string[]>;
type StoredDeckOverrides = Record<string, { deleted?: boolean; name?: string }>;
type ReviewMemorySource = "study" | "quiz" | "review";
type StoredReviewSchedule = Record<
  string,
  Record<
    string,
    {
      stage: number;
      ease: number;
      dueAt: string;
      lastReviewedAt?: string;
      lapses: number;
      correctStreak: number;
      source: ReviewMemorySource;
      lastRating: string;
    }
  >
>;
type SpeechErrorWindow = Window & {
  __wordloopShowSpeechError?: (message?: string) => void;
};

type TranslationReviewRow = {
  word: string;
  meaning: string;
  status: string;
  statusType: "correct" | "partial" | "review";
  reason: string;
};

type TranslationEvaluationResult = {
  score: number;
  source: string;
  critique: string;
  userTranslation: string;
  reference: string;
  reviewRows: TranslationReviewRow[];
};

type TranslationPromptData = {
  source: string;
  highlightedWords: string[];
};

type WordAiDetails = {
  content: string;
  etymology?: string;
  rootBreakdown?: string[];
  coreMeaning?: string;
  collocations: {
    phrase: string;
    translation: string;
  }[];
  derivatives?: {
    word: string;
    meaning: string;
  }[];
  synonymAnalysis?: {
    word: string;
    difference: string;
  }[];
  practicalPhrases: {
    phrase: string;
    translation: string;
  }[];
};

type WordDetailsState = {
  data?: WordAiDetails;
  error?: string;
  loading?: boolean;
};

type ImportedWordInput = {
  word: string;
  phonetic?: string;
  meaning?: string;
  example?: string;
  exampleTranslation?: string;
  synonyms?: string[];
};

type PendingDeckConfirm = {
  action: DeckConfirmAction;
  deck: VocabularyDeck;
};

const progressStorageKey = "wordloop.progress.v1";
const settingsStorageKey = "wordloop.settings.v1";
const masteryStorageKey = "wordloop.mastery.v1";
const reviewCyclesStorageKey = "wordloop.reviewCycles.v1";
const reviewStatsStorageKey = "wordloop.reviewStats.v1";
const dailyReviewedWordsStorageKey = "wordloop.dailyReviewedWords.v1";
const userDecksStorageKey = "wordloop.userDecks.v1";
const deckOverridesStorageKey = "wordloop.deckOverrides.v1";
const reviewScheduleStorageKey = "wordloop.reviewSchedule.v1";
const activeStorageUserKey = "wordloop.activeUserId.v1";
const defaultSettings: AppSettings = {
  deepseekApiKey: "",
  dailyStudyCount: 20,
  dailyReviewGoal: 30,
};

const review: ReviewData = {
  state: "notStarted",
  status: "尚未开始复习",
  percent: 0,
  due: 50,
  overdue: 12,
  completed: 0,
};

const reviewSessionLimit = 30;
const masteryReviewCycleTarget = 5;
const masteryReviewStageTarget = 8;
const reviewStageIntervalsMinutes = [
  5,
  30,
  12 * 60,
  24 * 60,
  2 * 24 * 60,
  4 * 24 * 60,
  7 * 24 * 60,
  15 * 24 * 60,
  30 * 24 * 60,
  60 * 24 * 60,
];

function getActiveStorageUserId() {
  if (typeof window === "undefined") return "";
  return window.localStorage.getItem(activeStorageUserKey) ?? "";
}

function setActiveStorageUserId(userId: string | null) {
  if (typeof window === "undefined") return;
  if (userId) window.localStorage.setItem(activeStorageUserKey, userId);
  else window.localStorage.removeItem(activeStorageUserKey);
}

function getScopedStorageKey(key: string, userId = getActiveStorageUserId()) {
  return userId ? `${key}.${userId}` : key;
}

function hasMeaningfulStoredValue(value: unknown): boolean {
  if (Array.isArray(value)) return value.length > 0;
  if (value && typeof value === "object") {
    return Object.values(value as Record<string, unknown>).some(hasMeaningfulStoredValue);
  }
  if (typeof value === "string") return value.trim().length > 0;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "boolean") return value;
  return false;
}

function parseStoredValue(raw: string | null) {
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function readStorageRaw(key: string) {
  if (typeof window === "undefined") return null;
  return window.localStorage.getItem(getScopedStorageKey(key));
}

function copyLegacyStorageToCurrentUser(keys: string[]) {
  if (typeof window === "undefined") return false;
  let copied = false;
  keys.forEach((key) => {
    const scopedKey = getScopedStorageKey(key);
    if (scopedKey === key || window.localStorage.getItem(scopedKey)) return;
    const legacyRaw = window.localStorage.getItem(key);
    if (!hasMeaningfulStoredValue(parseStoredValue(legacyRaw))) return;
    window.localStorage.setItem(scopedKey, legacyRaw as string);
    copied = true;
  });
  return copied;
}

function getLearningCalendarDays(hasLearningRecord: boolean): CalendarDay[] {
  const today = new Date();
  return Array.from({ length: 21 }, (_, index) => {
    const date = new Date(today);
    date.setDate(today.getDate() - (20 - index));
    const isToday = date.toDateString() === today.toDateString();
    return {
      label: String(date.getDate()),
      status: isToday ? (hasLearningRecord ? "done" : "today") : "muted",
    };
  });
}

type RecordStat = {
  label: string;
  value: string;
  suffix: string;
};

function addMinutes(date: Date, minutes: number) {
  return new Date(date.getTime() + minutes * 60 * 1000);
}

function getDayStart(date = new Date()) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function clampReviewStage(stage: number) {
  return Math.max(0, Math.min(reviewStageIntervalsMinutes.length - 1, stage));
}

function getReviewIntervalMinutes(stage: number, ease: number) {
  const baseMinutes = reviewStageIntervalsMinutes[clampReviewStage(stage)];
  return Math.max(5, Math.round(baseMinutes * Math.max(0.6, Math.min(1.8, ease))));
}

function createReviewMemory(stage: number, ease: number, source: ReviewMemorySource, lastRating: string, now = new Date()) {
  const safeStage = clampReviewStage(stage);
  return {
    stage: safeStage,
    ease: Math.max(0.7, Math.min(1.8, ease)),
    dueAt: addMinutes(now, getReviewIntervalMinutes(safeStage, ease)).toISOString(),
    lastReviewedAt: now.toISOString(),
    lapses: 0,
    correctStreak: safeStage > 0 ? 1 : 0,
    source,
    lastRating,
  };
}

function getFallbackReviewMemory(status: MarkStatus | undefined, reviewCycle = 0): StoredReviewSchedule[string][string] {
  const stageFromStatus = status === "known" ? 2 : status === "familiar" ? 1 : 0;
  const stage = Math.max(stageFromStatus, reviewCycle);
  return {
    stage,
    ease: status === "known" ? 1.15 : status === "familiar" ? 1 : 0.85,
    dueAt: new Date(0).toISOString(),
    lapses: 0,
    correctStreak: reviewCycle,
    source: "study",
    lastRating: status ?? "legacy",
  };
}

function getWordMemory(
  deckId: string,
  wordId: string,
  progress: StoredProgress,
  reviewCycles: StoredReviewCycles,
  reviewSchedule: StoredReviewSchedule,
) {
  return (
    reviewSchedule[deckId]?.[wordId] ??
    getFallbackReviewMemory(progress[deckId]?.[wordId], reviewCycles[deckId]?.[wordId] ?? 0)
  );
}

function isReviewDue(memory: StoredReviewSchedule[string][string], now = new Date()) {
  return new Date(memory.dueAt).getTime() <= now.getTime();
}

function isReviewOverdue(memory: StoredReviewSchedule[string][string], now = new Date()) {
  return new Date(memory.dueAt).getTime() < getDayStart(now).getTime();
}

function isMemoryMastered(memory: StoredReviewSchedule[string][string]) {
  return memory.stage >= masteryReviewStageTarget && memory.correctStreak >= masteryReviewStageTarget;
}

function scheduleAfterStudyAndQuiz(
  studyStatus: MarkStatus,
  quizStatusType: TranslationReviewRow["statusType"] | undefined,
  existingMemory?: StoredReviewSchedule[string][string],
) {
  const baseStage = studyStatus === "known" ? 2 : studyStatus === "familiar" ? 1 : 0;
  const quizDelta = quizStatusType === "correct" ? 2 : quizStatusType === "partial" ? 1 : quizStatusType === "review" ? -1 : 0;
  const baseEase = studyStatus === "known" ? 1.15 : studyStatus === "familiar" ? 1 : 0.85;
  const quizEase = quizStatusType === "correct" ? 0.08 : quizStatusType === "partial" ? -0.02 : quizStatusType === "review" ? -0.18 : 0;
  const nextStage = clampReviewStage(Math.max(existingMemory?.stage ?? 0, baseStage + quizDelta));
  const nextEase = Math.max(0.7, Math.min(1.8, (existingMemory?.ease ?? baseEase) + quizEase));
  const now = new Date();
  return {
    ...createReviewMemory(nextStage, nextEase, "quiz", `${studyStatus}:${quizStatusType ?? "ungraded"}`, now),
    lapses: existingMemory?.lapses ?? 0,
    correctStreak: quizStatusType === "review" ? 0 : Math.max(existingMemory?.correctStreak ?? 0, nextStage),
  };
}

function scheduleAfterReview(
  existingMemory: StoredReviewSchedule[string][string],
  status: MarkStatus,
) {
  const now = new Date();
  const nextStage =
    status === "known"
      ? existingMemory.stage + 2
      : status === "familiar"
        ? Math.max(1, existingMemory.stage + 1)
        : 0;
  const nextEase =
    status === "known"
      ? existingMemory.ease + 0.1
      : status === "familiar"
        ? existingMemory.ease - 0.04
        : existingMemory.ease - 0.22;
  return {
    ...createReviewMemory(nextStage, nextEase, "review", status, now),
    lapses: existingMemory.lapses + (status === "unknown" ? 1 : 0),
    correctStreak: status === "unknown" ? 0 : existingMemory.correctStreak + (status === "known" ? 2 : 1),
  };
}

function isWordMastered(
  deckId: string,
  wordId: string,
  mastery: StoredMastery,
  reviewCycles: StoredReviewCycles,
  reviewSchedule?: StoredReviewSchedule,
) {
  const memory = reviewSchedule?.[deckId]?.[wordId];
  if (memory) return isMemoryMastered(memory);
  return Boolean(mastery[deckId]?.[wordId] && (reviewCycles[deckId]?.[wordId] ?? 0) >= masteryReviewCycleTarget);
}

function getProgressSummary(
  progress: StoredProgress,
  mastery: StoredMastery,
  reviewCycles: StoredReviewCycles,
  reviewSchedule: StoredReviewSchedule,
  decks: VocabularyDeck[],
) {
  const deckWordIds = new Map(decks.map((deck) => [deck.id, new Set(deck.words.map((word) => word.id))]));
  const learnedEntries = decks.flatMap((deck) => {
    const wordIds = deckWordIds.get(deck.id) ?? new Set<string>();
    return Object.entries(progress[deck.id] ?? {})
      .filter(([wordId]) => wordIds.has(wordId))
      .map(([wordId, status]) => ({ deckId: deck.id, status, wordId }));
  });
  const learned = learnedEntries.length;
  const reviewing = learnedEntries.filter(({ deckId, wordId }) => {
    const memory = getWordMemory(deckId, wordId, progress, reviewCycles, reviewSchedule);
    return !isMemoryMastered(memory) && isReviewDue(memory);
  }).length;
  const overdue = learnedEntries.filter(({ deckId, wordId }) => {
    const memory = getWordMemory(deckId, wordId, progress, reviewCycles, reviewSchedule);
    return !isMemoryMastered(memory) && isReviewOverdue(memory);
  }).length;
  const mastered = learnedEntries.filter(({ deckId, wordId }) =>
    isWordMastered(deckId, wordId, mastery, reviewCycles, reviewSchedule),
  ).length;
  const totalWords = decks.reduce((sum, deck) => sum + deck.words.length, 0);
  const percent = learned ? Math.round((mastered / learned) * 100) : 0;

  return { learned, reviewing, overdue, mastered, totalWords, percent };
}

function getTodayKey() {
  const today = new Date();
  return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
}

function getHomeReviewData(
  progress: StoredProgress,
  mastery: StoredMastery,
  reviewCycles: StoredReviewCycles,
  reviewSchedule: StoredReviewSchedule,
  decks: VocabularyDeck[],
  dailyReviewGoal: number,
  reviewStats: StoredReviewStats,
): ReviewData {
  const summary = getProgressSummary(progress, mastery, reviewCycles, reviewSchedule, decks);
  const completedToday = reviewStats[getTodayKey()] ?? 0;
  const todayTarget = Math.min(dailyReviewGoal, completedToday + summary.reviewing);
  const percent = todayTarget > 0 ? Math.min(100, Math.round((completedToday / todayTarget) * 100)) : 0;
  const goalCompleted = todayTarget > 0 && completedToday >= todayTarget;
  const state: ReviewState = goalCompleted
    ? summary.reviewing > 0
      ? "todayDone"
      : "allDone"
    : completedToday > 0
      ? "inProgress"
      : "notStarted";

  return {
    state,
    status: state === "notStarted" ? "尚未开始复习" : goalCompleted ? "完成度已达标" : "复习中",
    percent,
    due: summary.reviewing,
    overdue: summary.overdue,
    completed: completedToday,
  };
}

function getLearningRecordStats(
  progress: StoredProgress,
  mastery: StoredMastery,
  reviewCycles: StoredReviewCycles,
  reviewSchedule: StoredReviewSchedule,
  decks: VocabularyDeck[],
): RecordStat[] {
  const summary = getProgressSummary(progress, mastery, reviewCycles, reviewSchedule, decks);
  return [
    { label: "连续打卡", value: summary.learned > 0 ? "1" : "0", suffix: "天" },
    { label: "累计学习", value: summary.learned > 0 ? "1" : "0", suffix: "天" },
    { label: "已学习", value: String(summary.learned), suffix: "词" },
    { label: "已掌握", value: String(summary.mastered), suffix: "词" },
  ];
}

function createEmptyProgress(): StoredProgress {
  return {
    designer: {},
    developer: {},
    office: {},
    travel: {},
  };
}

function createEmptyMastery(): StoredMastery {
  return {
    designer: {},
    developer: {},
    office: {},
    travel: {},
  };
}

function createEmptyReviewCycles(): StoredReviewCycles {
  return {
    designer: {},
    developer: {},
    office: {},
    travel: {},
  };
}

function loadProgress(): StoredProgress {
  if (typeof window === "undefined") return createEmptyProgress();

  try {
    const raw = readStorageRaw(progressStorageKey);
    if (!raw) return createEmptyProgress();
    return { ...createEmptyProgress(), ...JSON.parse(raw) };
  } catch {
    return createEmptyProgress();
  }
}

function saveProgress(progress: StoredProgress) {
  window.localStorage.setItem(getScopedStorageKey(progressStorageKey), JSON.stringify(progress));
}

function loadMastery(): StoredMastery {
  if (typeof window === "undefined") return createEmptyMastery();

  try {
    const raw = readStorageRaw(masteryStorageKey);
    if (!raw) return createEmptyMastery();
    return { ...createEmptyMastery(), ...JSON.parse(raw) };
  } catch {
    return createEmptyMastery();
  }
}

function saveMastery(mastery: StoredMastery) {
  window.localStorage.setItem(getScopedStorageKey(masteryStorageKey), JSON.stringify(mastery));
}

function loadReviewCycles(): StoredReviewCycles {
  if (typeof window === "undefined") return createEmptyReviewCycles();

  try {
    const raw = readStorageRaw(reviewCyclesStorageKey);
    if (!raw) return createEmptyReviewCycles();
    return { ...createEmptyReviewCycles(), ...JSON.parse(raw) };
  } catch {
    return createEmptyReviewCycles();
  }
}

function saveReviewCycles(reviewCycles: StoredReviewCycles) {
  window.localStorage.setItem(getScopedStorageKey(reviewCyclesStorageKey), JSON.stringify(reviewCycles));
}

function loadReviewStats(): StoredReviewStats {
  if (typeof window === "undefined") return {};

  try {
    const raw = readStorageRaw(reviewStatsStorageKey);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveReviewStats(reviewStats: StoredReviewStats) {
  window.localStorage.setItem(getScopedStorageKey(reviewStatsStorageKey), JSON.stringify(reviewStats));
}

function loadDailyReviewedWords(): StoredDailyReviewedWords {
  if (typeof window === "undefined") return {};

  try {
    const raw = readStorageRaw(dailyReviewedWordsStorageKey);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function saveDailyReviewedWords(dailyReviewedWords: StoredDailyReviewedWords) {
  window.localStorage.setItem(getScopedStorageKey(dailyReviewedWordsStorageKey), JSON.stringify(dailyReviewedWords));
}

function loadReviewSchedule(): StoredReviewSchedule {
  if (typeof window === "undefined") return {};

  try {
    const raw = readStorageRaw(reviewScheduleStorageKey);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function saveReviewSchedule(reviewSchedule: StoredReviewSchedule) {
  window.localStorage.setItem(getScopedStorageKey(reviewScheduleStorageKey), JSON.stringify(reviewSchedule));
}

function loadDeckOverrides(): StoredDeckOverrides {
  if (typeof window === "undefined") return {};

  try {
    const raw = readStorageRaw(deckOverridesStorageKey);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function saveDeckOverrides(deckOverrides: StoredDeckOverrides) {
  window.localStorage.setItem(getScopedStorageKey(deckOverridesStorageKey), JSON.stringify(deckOverrides));
}

function loadUserDecks(): VocabularyDeck[] {
  if (typeof window === "undefined") return [];

  try {
    const raw = readStorageRaw(userDecksStorageKey);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveUserDecks(decks: VocabularyDeck[]) {
  window.localStorage.setItem(getScopedStorageKey(userDecksStorageKey), JSON.stringify(decks));
}

function slugifyDeckName(name: string) {
  return name
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, "-")
    .replace(/^-|-$/g, "");
}

function getImportedWordText(input: string | ImportedWordInput) {
  return typeof input === "string" ? input.trim() : input.word.trim();
}

function createLocalWord(deckId: string, input: string | ImportedWordInput, index: number): VocabularyWord {
  const normalized = getImportedWordText(input);
  const data: Partial<ImportedWordInput> = typeof input === "string" ? {} : input;
  const meaning = data.meaning?.trim() || "待补全释义";
  const example = data.example?.trim() || `${normalized} is added to your vocabulary deck.`;
  const exampleTranslation = data.exampleTranslation?.trim() || `${normalized} 已加入你的词库。`;
  return {
    id: `${deckId}-${Date.now()}-${index}`,
    word: normalized,
    phonetic: data.phonetic?.trim() || `/${normalized.toLowerCase().replace(/[^a-z]+/g, " ").trim() || normalized}/`,
    meaning,
    example,
    exampleTranslation,
    synonyms: data.synonyms ?? [],
    extraMeanings: [meaning],
    extraExamples: [
      {
        sentence: example,
        translation: exampleTranslation,
      },
    ],
  };
}

function parseWordInput(input: string) {
  return Array.from(
    new Set(
      input
        .split(/[\s,，;；、\n\r]+/)
        .map((word) => word.trim())
        .filter(Boolean),
    ),
  );
}

async function parseWordsFromFile(file: File) {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (extension === "xlsx" || extension === "xls") {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const rows = workbook.SheetNames.flatMap((sheetName) =>
      XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, blankrows: false }),
    );
    return parseWordRows(rows);
  }
  return parseTextWordRows(await file.text());
}

function parseTextWordRows(text: string): ImportedWordInput[] {
  const lines = text
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  const rowLike = lines.some((line) => /,|\t/.test(line));
  if (!rowLike) return parseWordInput(text).map((word) => ({ word }));
  return parseWordRows(lines.map((line) => line.split(/\t|,/)));
}

function parseWordRows(rows: unknown[][]): ImportedWordInput[] {
  let wordColumn = 0;
  let meaningColumn = 1;
  const firstRow = rows.find((row) => row.some((cell) => String(cell ?? "").trim()));
  if (firstRow) {
    const headers = firstRow.map((cell) => String(cell ?? "").trim().toLowerCase());
    const detectedWordColumn = headers.findIndex((header) => ["单词", "word", "vocabulary", "term"].includes(header));
    const detectedMeaningColumn = headers.findIndex((header) => ["释义", "中文释义", "meaning", "definition"].includes(header));
    if (detectedWordColumn >= 0) wordColumn = detectedWordColumn;
    if (detectedMeaningColumn >= 0) meaningColumn = detectedMeaningColumn;
  }

  const seen = new Set<string>();
  const junk = new Set(["word", "meaning", "definition", "term", "vocabulary", "单词", "释义", "head", "body", "html", "table", "thead", "tbody", "tr", "td", "th"]);

  return rows
    .map((row) => {
      const word = String(row[wordColumn] ?? "").trim();
      const meaning = String(row[meaningColumn] ?? "").trim();
      return { word, meaning };
    })
    .filter((item) => {
      const normalized = item.word.toLowerCase();
      if (!normalized || seen.has(normalized) || junk.has(normalized)) return false;
      if (/[\u4e00-\u9fa5]/.test(item.word)) return false;
      if (!/[a-z]/i.test(item.word)) return false;
      if (item.word.length > 48) return false;
      seen.add(normalized);
      return true;
    });
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getExportFileName(deckName: string) {
  const safeName = deckName.trim().replace(/[\\/:*?"<>|]+/g, "-") || "词库";
  return `${safeName}.xls`;
}

function downloadDeckAsExcel(deck: VocabularyDeck) {
  if (typeof window === "undefined") return;
  const rows = deck.words
    .map(
      (word) => `
        <tr>
          <td>${escapeHtml(word.word)}</td>
          <td>${escapeHtml(word.meaning)}</td>
        </tr>
      `,
    )
    .join("");
  const table = `
    <html>
      <head><meta charset="UTF-8" /></head>
      <body>
        <table>
          <thead>
            <tr>
              <th>单词</th>
              <th>释义</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </body>
    </html>
  `;
  const blob = new Blob([table], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = getExportFileName(deck.name);
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.URL.revokeObjectURL(url);
}

function canUseSpeechSynthesis() {
  return typeof window !== "undefined" && "speechSynthesis" in window && "SpeechSynthesisUtterance" in window;
}

function speakText(text: string, options: { lang?: "en-US" | "zh-CN"; notify?: boolean } = {}) {
  if (!canUseSpeechSynthesis()) {
    if (options.notify && typeof window !== "undefined") {
      (window as SpeechErrorWindow).__wordloopShowSpeechError?.("当前浏览器不支持系统语音播放，请用 Chrome 打开后重试");
    }
    return false;
  }
  const cleanText = text.trim();
  if (!cleanText) return false;
  window.speechSynthesis.cancel();
  window.speechSynthesis.resume();

  function createUtterance() {
    const nextUtterance = new window.SpeechSynthesisUtterance(cleanText);
    nextUtterance.lang = options.lang ?? (/^[a-zA-Z\s'-]+$/.test(cleanText) ? "en-US" : "zh-CN");
    const preferredVoice = getPreferredTingtingVoice(nextUtterance.lang);
    if (preferredVoice) nextUtterance.voice = preferredVoice;
    nextUtterance.rate = 0.92;
    nextUtterance.pitch = 1;
    nextUtterance.volume = 1;
    return nextUtterance;
  }

  const utterance = createUtterance();
  let started = false;
  utterance.onstart = () => {
    started = true;
  };
  utterance.lang = options.lang ?? (/^[a-zA-Z\s'-]+$/.test(cleanText) ? "en-US" : "zh-CN");
  window.speechSynthesis.speak(utterance);
  window.setTimeout(() => {
    if (started || window.speechSynthesis.speaking || window.speechSynthesis.pending) return;
    window.speechSynthesis.cancel();
    window.speechSynthesis.resume();
    window.speechSynthesis.speak(createUtterance());
  }, 180);
  if (options.notify) {
    window.setTimeout(() => {
      if (started || window.speechSynthesis.speaking) return;
      (window as SpeechErrorWindow).__wordloopShowSpeechError?.("语音播放未启动，请确认浏览器支持语音播放且页面未被静音。");
    }, 1000);
  }
  return true;
}

function getPreferredTingtingVoice(lang: string) {
  if (typeof window === "undefined" || !("speechSynthesis" in window)) return null;
  const voices = window.speechSynthesis.getVoices();
  if (!voices.length) {
    window.speechSynthesis.onvoiceschanged = () => window.speechSynthesis.getVoices();
    return null;
  }
  const normalizedLang = lang.toLowerCase();
  if (normalizedLang.startsWith("en")) {
    return (
      voices.find((voice) => voice.lang.toLowerCase().startsWith("en") && /samantha|ava|allison|google us english/i.test(voice.name)) ??
      voices.find((voice) => voice.lang.toLowerCase().startsWith("en")) ??
      null
    );
  }
  return (
    voices.find((voice) => /婷婷|tingting/i.test(voice.name)) ??
    voices.find((voice) => voice.lang.toLowerCase().startsWith("zh-cn")) ??
    voices.find((voice) => voice.lang.toLowerCase().startsWith(normalizedLang.slice(0, 2))) ??
    null
  );
}

function useDelayedSpeech(text: string, delay = 650, options: { lang?: "en-US" | "zh-CN" } = {}) {
  const timerRef = useRef<number | null>(null);

  useEffect(() => {
    return () => {
      if (timerRef.current) window.clearTimeout(timerRef.current);
    };
  }, []);

  function clearSpeechTimer() {
    if (!timerRef.current) return;
    window.clearTimeout(timerRef.current);
    timerRef.current = null;
  }

  return {
    onPointerEnter: () => {
      clearSpeechTimer();
      timerRef.current = window.setTimeout(() => {
        speakText(text, options);
        timerRef.current = null;
      }, delay);
    },
    onPointerLeave: clearSpeechTimer,
  };
}

function getSpeechRecognitionConstructor() {
  if (typeof window === "undefined") return null;
  const speechWindow = window as Window &
    typeof globalThis & {
      SpeechRecognition?: new () => SpeechRecognitionLike;
      webkitSpeechRecognition?: new () => SpeechRecognitionLike;
    };
  return speechWindow.SpeechRecognition ?? speechWindow.webkitSpeechRecognition ?? null;
}

type SpeechRecognitionLike = {
  continuous: boolean;
  interimResults: boolean;
  lang: string;
  onend: (() => void) | null;
  onerror: ((event: { error?: string }) => void) | null;
  onresult: ((event: { results: ArrayLike<{ 0: { transcript: string } }> }) => void) | null;
  start: () => void;
  stop: () => void;
};

function loadSettings(): AppSettings {
  if (typeof window === "undefined") return defaultSettings;

  try {
    const raw = readStorageRaw(settingsStorageKey);
    if (!raw) return defaultSettings;
    const scopedSettings = { ...defaultSettings, ...JSON.parse(raw) };
    const legacyRaw = getScopedStorageKey(settingsStorageKey) === settingsStorageKey
      ? null
      : window.localStorage.getItem(settingsStorageKey);
    if (!scopedSettings.deepseekApiKey.trim() && legacyRaw) {
      const legacySettings = JSON.parse(legacyRaw) as Partial<AppSettings>;
      if (legacySettings.deepseekApiKey?.trim()) {
        return { ...scopedSettings, deepseekApiKey: legacySettings.deepseekApiKey };
      }
    }
    return scopedSettings;
  } catch {
    return defaultSettings;
  }
}

function saveSettings(settings: AppSettings) {
  window.localStorage.setItem(getScopedStorageKey(settingsStorageKey), JSON.stringify(settings));
}

function extractJsonObject(content: string) {
  const fenced = content.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const raw = fenced?.[1] ?? content;
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  if (start < 0 || end < start) throw new Error("AI 返回内容不是 JSON。");
  return JSON.parse(raw.slice(start, end + 1));
}

async function requestDeepSeekJson(apiKey: string, prompt: string) {
  const response = await fetch("/api/deepseek/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "deepseek-chat",
      temperature: 0.2,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content: "你是英语学习产品中的翻译测评助手。只输出合法 JSON，不要输出 Markdown。",
        },
        { role: "user", content: prompt },
      ],
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || `DeepSeek 请求失败：${response.status}`);
  }

  const payload = await response.json();
  const content = payload?.choices?.[0]?.message?.content;
  if (typeof content !== "string") throw new Error("DeepSeek 返回格式异常。");
  return extractJsonObject(content);
}

async function requestDeepSeekText(apiKey: string, prompt: string) {
  const response = await fetch("/api/deepseek/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "deepseek-chat",
      temperature: 0.25,
      messages: [
        {
          role: "system",
          content: "你是英语学习产品中的词汇讲解助手。输出简洁中文正文，可以保留少量英文短语，不要输出 JSON。",
        },
        { role: "user", content: prompt },
      ],
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText || `DeepSeek 请求失败：${response.status}`);
  }

  const payload = await response.json();
  const content = payload?.choices?.[0]?.message?.content;
  if (typeof content !== "string") throw new Error("DeepSeek 返回格式异常。");
  return content.trim();
}

function normalizeEvaluationResult(
  data: any,
  source: string,
  userTranslation: string,
  targetWords: string[],
): TranslationEvaluationResult {
  const rows = Array.isArray(data.wordResults) ? data.wordResults : [];
  const normalizedRows: TranslationReviewRow[] = rows.map((row: any) => {
    const partial = row.level === "partial" || row.level === "partially_correct" || row.statusType === "partial";
    const weak = row.level === "weak" || row.level === "review" || row.statusType === "review";
    return {
      word: String(row.word || "—"),
      meaning: String(row.suggestedTranslation || row.meaning || "—"),
      status: weak ? "待复习" : partial ? "基本理解" : "正确",
      statusType: weak ? "review" : partial ? "partial" : "correct",
      reason: String(row.comment || row.reason || "—"),
    } satisfies TranslationReviewRow;
  });
  const rowByWord = new Map(normalizedRows.map((row) => [row.word.trim().toLowerCase(), row]));
  const mergedRows = targetWords.length
    ? targetWords.map((word) => {
        const matched = rowByWord.get(word.trim().toLowerCase());
        return (
          matched ?? {
            word,
            meaning: "请回到词卡复习释义",
            status: "待复习",
            statusType: "review" as const,
            reason: "AI 测评未返回该词的单独分析，已保留在本次复盘表中。",
          }
        );
      })
    : normalizedRows;
  return {
    score: Math.max(0, Math.min(100, Number(data.score) || 0)),
    source,
    critique: String(data.overallComment || data.critique || "已完成翻译测评。"),
    userTranslation,
    reference: String(data.referenceTranslation || data.reference || "DeepSeek 未返回参考译文。"),
    reviewRows: mergedRows,
  };
}

function normalizeWordDetails(data: any): WordAiDetails {
  if (typeof data === "string") {
    return {
      content: data,
      collocations: [],
      practicalPhrases: [],
    };
  }

  if (typeof data?.content === "string") {
    return {
      ...data,
      content: data.content,
      collocations: Array.isArray(data.collocations) ? data.collocations : [],
      practicalPhrases: Array.isArray(data.practicalPhrases) ? data.practicalPhrases : [],
    };
  }

  const rootBreakdown = Array.isArray(data.rootBreakdown)
    ? data.rootBreakdown.map((item: unknown) => String(item)).filter(Boolean)
    : [];
  const collocations: WordAiDetails["collocations"] = Array.isArray(data.collocations)
    ? data.collocations.slice(0, 5).map((item: any) => ({
        phrase: String(item.phrase || item.example || ""),
        translation: String(item.translation || item.meaning || ""),
      }))
    : [];
  const derivatives: NonNullable<WordAiDetails["derivatives"]> = Array.isArray(data.derivatives)
    ? data.derivatives.slice(0, 6).map((item: any) => ({
        word: String(item.word || ""),
        meaning: String(item.meaning || ""),
      }))
    : [];
  const synonymAnalysis: NonNullable<WordAiDetails["synonymAnalysis"]> = Array.isArray(data.synonymAnalysis)
    ? data.synonymAnalysis.slice(0, 5).map((item: any) => ({
        word: String(item.word || ""),
        difference: String(item.difference || item.explanation || ""),
      }))
    : [];
  const practicalPhrases: WordAiDetails["practicalPhrases"] = Array.isArray(data.practicalPhrases)
    ? data.practicalPhrases.slice(0, 6).map((item: any) => ({
        phrase: String(item.phrase || ""),
        translation: String(item.translation || item.meaning || ""),
      }))
    : [];
  const content = [
    data.etymology || rootBreakdown.length ? `词根词源拆解\n${[data.etymology, ...rootBreakdown].filter(Boolean).join(" / ")}` : "",
    `核心释义\n${String(data.coreMeaning || data.meaning || "")}`,
    collocations.length || practicalPhrases.length
      ? `搭配短句和实用例句\n${[...collocations, ...practicalPhrases]
          .map((item) => `${item.phrase}：${item.translation}`)
          .join("\n")}`
      : "",
    derivatives.length || synonymAnalysis.length
      ? `衍生词和近义词辨析\n${[
          ...derivatives.map((item) => `${item.word}：${item.meaning}`),
          ...synonymAnalysis.map((item) => `${item.word}：${item.difference}`),
        ].join("\n")}`
      : "",
  ]
    .filter(Boolean)
    .join("\n\n");
  return {
    content,
    etymology: data.etymology ? String(data.etymology) : undefined,
    rootBreakdown,
    coreMeaning: String(data.coreMeaning || data.meaning || ""),
    collocations,
    derivatives,
    synonymAnalysis,
    practicalPhrases,
  };
}

async function generateWordDetailsWithDeepSeek(apiKey: string, word: VocabularyWord) {
  const prompt = `
请为英语学习应用生成单词详情，面向中文母语的设计/职场英语学习者。

单词: ${word.word}
基础释义: ${word.meaning}
原例句: ${word.example}

要求:
- 只输出纯文本，不要 JSON，不要代码块
- 用以下四个小标题，标题独占一行：词根词源拆解、核心释义、搭配短句和实用例句、衍生词和近义词辨析
- 每个部分 1-3 行，保持紧凑
- “搭配短句和实用例句”不要重复原例句，不要输出与原例句意思高度相同的句子
- “搭配短句和实用例句”里的每条英文后必须附中文翻译，格式建议为：英文短句：中文翻译
- 英文短句要自然、实用，适合用户 hover 后朗读
- 如果没有明确词源、衍生词或近义词辨析，就写“暂无可靠补充”
- 中文解释要简洁
`;
  const content = await requestDeepSeekText(apiKey, prompt);
  return normalizeWordDetails(content);
}

async function generateDeckWordsWithDeepSeek(apiKey: string, topic: string, count: number) {
  const prompt = `
请为英语单词学习应用生成一个主题词库。

主题：${topic}
数量：${count}

要求：
- 只输出 JSON
- words 必须是英文单词或常用英文短语
- 尽量贴合主题，避免重复，避免过长短语
- 每个词必须包含中文释义、音标、自然英文例句、中文例句翻译
- synonyms 可为空数组，但不要编造不相关近义词

JSON schema:
{
  "words": [
    {
      "word": "english word or phrase",
      "phonetic": "/phonetic/",
      "meaning": "中文释义",
      "example": "Natural English example sentence.",
      "exampleTranslation": "中文例句翻译",
      "synonyms": ["synonym"]
    }
  ]
}
`;
  const data = await requestDeepSeekJson(apiKey, prompt);
  if (!Array.isArray(data.words)) throw new Error("AI 未返回可导入的词表。");
  return data.words
    .map((item: any) =>
      typeof item === "string"
        ? { word: item }
        : {
            word: String(item.word || "").trim(),
            phonetic: String(item.phonetic || "").trim(),
            meaning: String(item.meaning || "").trim(),
            example: String(item.example || "").trim(),
            exampleTranslation: String(item.exampleTranslation || item.translation || "").trim(),
            synonyms: Array.isArray(item.synonyms) ? item.synonyms.map((word: unknown) => String(word).trim()).filter(Boolean) : [],
          },
    )
    .filter((item: ImportedWordInput) => item.word)
    .slice(0, count);
}

function needsWordCompletion(word: string | ImportedWordInput) {
  if (typeof word === "string") return true;
  return !word.meaning?.trim() || !word.example?.trim() || !word.exampleTranslation?.trim();
}

async function completeImportedWordsWithDeepSeek(apiKey: string, words: Array<string | ImportedWordInput>) {
  const result: ImportedWordInput[] = words.map((word) => (typeof word === "string" ? { word } : word));
  const incompleteIndexes = result
    .map((word, index) => ({ word, index }))
    .filter(({ word }) => needsWordCompletion(word))
    .map(({ index }) => index);

  if (!incompleteIndexes.length) return result;
  if (!apiKey.trim()) return result;

  const batchSize = 40;
  for (let start = 0; start < incompleteIndexes.length; start += batchSize) {
    const indexes = incompleteIndexes.slice(start, start + batchSize);
    const targetWords = indexes.map((index) => result[index].word);
    const prompt = `
请为英语单词学习应用补齐导入词条信息。

目标词：
${targetWords.map((word) => `- ${word}`).join("\n")}

要求：
- 只输出 JSON
- 每个输入词都必须返回
- meaning 使用中文简洁释义
- example 是自然英文例句
- exampleTranslation 是例句中文翻译
- phonetic 尽量给出音标；无法确定可为空字符串
- synonyms 可为空数组

JSON schema:
{
  "words": [
    {
      "word": "原单词",
      "phonetic": "/phonetic/",
      "meaning": "中文释义",
      "example": "Natural English example sentence.",
      "exampleTranslation": "中文例句翻译",
      "synonyms": ["synonym"]
    }
  ]
}
`;
    const data = await requestDeepSeekJson(apiKey.trim(), prompt);
    if (!Array.isArray(data.words)) throw new Error("AI 未返回可补齐的词条。");
    const completions = new Map<string, ImportedWordInput>(
      data.words.map((item: any) => [
        String(item.word || "").trim().toLowerCase(),
        {
          word: String(item.word || "").trim(),
          phonetic: String(item.phonetic || "").trim(),
          meaning: String(item.meaning || "").trim(),
          example: String(item.example || "").trim(),
          exampleTranslation: String(item.exampleTranslation || item.translation || "").trim(),
          synonyms: Array.isArray(item.synonyms) ? item.synonyms.map((word: unknown) => String(word).trim()).filter(Boolean) : [],
        } satisfies ImportedWordInput,
      ]),
    );

    indexes.forEach((index) => {
      const current = result[index];
      const completed = completions.get(current.word.trim().toLowerCase());
      if (!completed) return;
      result[index] = {
        ...current,
        phonetic: current.phonetic || completed.phonetic,
        meaning: current.meaning || completed.meaning,
        example: current.example || completed.example,
        exampleTranslation: current.exampleTranslation || completed.exampleTranslation,
        synonyms: current.synonyms?.length ? current.synonyms : completed.synonyms,
      };
    });
  }

  return result;
}

function getWordDetailKey(deckId: string, wordId: string) {
  return `${deckId}:${wordId}`;
}

async function loadCachedWordDetails(userId: string, deckId: string, wordId: string) {
  const { data, error } = await supabase
    .from("word_ai_details")
    .select("details")
    .eq("user_id", userId)
    .eq("word_key", getWordDetailKey(deckId, wordId))
    .maybeSingle();
  if (error) throw error;
  return data?.details ? normalizeWordDetails(data.details) : null;
}

async function saveCachedWordDetails(userId: string, deckId: string, wordId: string, details: WordAiDetails) {
  const { error } = await supabase.from("word_ai_details").upsert(
    {
      user_id: userId,
      word_key: getWordDetailKey(deckId, wordId),
      deck_local_id: deckId,
      word_local_id: wordId,
      details,
    },
    { onConflict: "user_id,word_key" },
  );
  if (error) throw error;
}

async function evaluateTranslationWithDeepSeek({
  apiKey,
  mode,
  source,
  userTranslation,
  targetWords,
}: {
  apiKey: string;
  mode: TranslationMode;
  source: string;
  userTranslation: string;
  targetWords: string[];
}) {
  const prompt = `
请评判用户的段落翻译，并返回严格 JSON。

翻译方向：${mode === "enToZh" ? "英译中" : "中译英"}
原文：
${source}

用户翻译：
${userTranslation}

目标词/表达：
${targetWords.join(", ")}

JSON schema:
{
  "score": 0-100,
  "overallComment": "中文点评，说明整体表现和主要问题",
  "referenceTranslation": "参考译文",
  "wordResults": [
    {
      "word": "目标词",
      "level": "correct、partial 或 weak",
      "suggestedTranslation": "建议译法或中文释义",
      "comment": "中文原因解析。correct=答对；partial=不完全对、语义基本接近但用法/语境有偏差；weak=答错或遗漏"
    }
  ]
}
`;
  const data = await requestDeepSeekJson(apiKey, prompt);
  return normalizeEvaluationResult(data, source, userTranslation, targetWords);
}

async function generateTranslationPromptWithDeepSeek({
  apiKey,
  mode,
  targetWords,
}: {
  apiKey: string;
  mode: TranslationMode;
  targetWords: VocabularyWord[];
}): Promise<TranslationPromptData> {
  const englishWordTarget = Math.min(360, Math.max(120, targetWords.length * 16));
  const chineseCharTarget = Math.min(520, Math.max(180, targetWords.length * 24));
  const wordList = targetWords
    .map((word) => `- ${word.word}: ${word.meaning}; example: ${word.example}`)
    .join("\n");
  const prompt = `
请根据目标词生成一个段落翻译测验，并返回严格 JSON。

翻译方向：${mode === "enToZh" ? "英译中：生成英文原文，用户会翻译成中文" : "中译英：生成中文原文，用户会翻译成英文"}
目标词：
${wordList}

要求：
- ${mode === "enToZh" ? `英文原文约 ${englishWordTarget} 词` : `中文原文约 ${chineseCharTarget} 字`}
- 必须覆盖全部目标词。英译中时，每个目标词必须在英文 source 中至少出现一次；中译英时，每个目标词的中文语义必须在 source 中至少出现一次
- 内容自然，不要列清单
- highlightedWords 必须是 source 中真实出现、需要高亮的词或短语

JSON schema:
{
  "source": "用于展示给用户翻译的原文段落",
  "highlightedWords": ["source 中真实出现的词或短语"]
}
`;
  const data = await requestDeepSeekJson(apiKey, prompt);
  const source = String(data.source || "").trim();
  const highlightedWords = Array.isArray(data.highlightedWords)
    ? data.highlightedWords.map((word: unknown) => String(word).trim()).filter(Boolean)
    : targetWords.map((word) => word.word);

  if (!source) throw new Error("DeepSeek 未返回测验段落。");
  const mergedHighlights = Array.from(new Set([...highlightedWords, ...targetWords.map((word) => word.word)]));
  return { source, highlightedWords: mergedHighlights };
}

function getInitialStudyDeckId(): BuiltinDeckId | null {
  if (typeof window === "undefined") return null;
  const match = window.location.hash.match(/^#\/study\/(designer|developer|office|travel)$/);
  return match ? (match[1] as BuiltinDeckId) : null;
}

function getInitialView(): View {
  if (typeof window === "undefined") return "home";
  if (window.location.hash === "#/ai-translate") return "translate";
  if (window.location.hash === "#/ai-translate/result") return "translateResult";
  if (window.location.hash === "#/review") return "review";
  return getInitialStudyDeckId() ? "study" : "home";
}

function createSession(deck: VocabularyDeck, progress: StoredProgress, wordCount = defaultSettings.dailyStudyCount): StudySession {
  const deckProgress = progress[deck.id] ?? {};
  const unlearnedWords = deck.words.filter((word) => !deckProgress[word.id]);

  return {
    id: `${deck.id}-${Date.now()}`,
    deckId: deck.id,
    wordIds: unlearnedWords.slice(0, wordCount).map((word) => word.id),
    statuses: {},
    createdAt: new Date().toISOString(),
  };
}

function getDeckIdForWord(wordId: string, decks: VocabularyDeck[]): string | null {
  return decks.find((deck) => deck.words.some((word) => word.id === wordId))?.id ?? null;
}

function App() {
  const [settings, setSettings] = useState<AppSettings>(() => loadSettings());
  const [progress, setProgress] = useState<StoredProgress>(() => loadProgress());
  const [mastery, setMastery] = useState<StoredMastery>(() => loadMastery());
  const [reviewCycles, setReviewCycles] = useState<StoredReviewCycles>(() => loadReviewCycles());
  const [reviewStats, setReviewStats] = useState<StoredReviewStats>(() => loadReviewStats());
  const [dailyReviewedWords, setDailyReviewedWords] = useState<StoredDailyReviewedWords>(() => loadDailyReviewedWords());
  const [reviewSchedule, setReviewSchedule] = useState<StoredReviewSchedule>(() => loadReviewSchedule());
  const [userDecks, setUserDecks] = useState<VocabularyDeck[]>(() => loadUserDecks());
  const [deckOverrides, setDeckOverrides] = useState<StoredDeckOverrides>(() => loadDeckOverrides());
  const [session, setSession] = useState<StudySession | null>(() => {
    const initialDeckId = getInitialStudyDeckId();
    const initialDeck = initialDeckId ? getDeckById(initialDeckId) : undefined;
    return initialDeck ? createSession(initialDeck, loadProgress(), loadSettings().dailyStudyCount) : null;
  });
  const [view, setView] = useState<View>(() => getInitialView());
  const [expandedWordIds, setExpandedWordIds] = useState<Set<string>>(() => new Set());
  const [wordDetailsById, setWordDetailsById] = useState<Record<string, WordDetailsState>>({});
  const [quizNotice, setQuizNotice] = useState("");
  const [translationInput, setTranslationInput] = useState("");
  const [translationMode, setTranslationMode] = useState<TranslationMode>("enToZh");
  const [translationPrompt, setTranslationPrompt] = useState<TranslationPromptData | null>(null);
  const [translationTargetWords, setTranslationTargetWords] = useState<string[]>([]);
  const [translationStatus, setTranslationStatus] = useState<TranslationStatus>("idle");
  const [translationError, setTranslationError] = useState("");
  const [translationResult, setTranslationResult] = useState<TranslationEvaluationResult | null>(null);
  const [reviewDeckFilter, setReviewDeckFilter] = useState<ReviewDeckFilter>("all");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [quickModalOpen, setQuickModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [importTab, setImportTab] = useState<ImportDeckTab>("ai");
  const [importWordsDeckId, setImportWordsDeckId] = useState<string | null>(null);
  const [openDeckMenuId, setOpenDeckMenuId] = useState<string | null>(null);
  const [renamingDeckId, setRenamingDeckId] = useState<string | null>(null);
  const [expandingDeckId, setExpandingDeckId] = useState<string | null>(null);
  const [pendingDeckConfirm, setPendingDeckConfirm] = useState<PendingDeckConfirm | null>(null);
  const [toastMessage, setToastMessage] = useState("");
  const [speechSupported, setSpeechSupported] = useState(() => canUseSpeechSynthesis());
  const [authReady, setAuthReady] = useState(false);
  const [checkingAuth, setCheckingAuth] = useState(true);
  const [currentUserEmail, setCurrentUserEmail] = useState("");
  const supabaseUserIdRef = useRef<string | null>(null);
  const supabaseSyncTimerRef = useRef<number | null>(null);

  const allDecks: VocabularyDeck[] = useMemo(
    () => [
      ...builtinDecks
        .filter((deck) => !deckOverrides[deck.id]?.deleted)
        .map((deck) => ({ ...deck, name: deckOverrides[deck.id]?.name ?? deck.name })),
      ...userDecks,
    ],
    [deckOverrides, userDecks],
  );
  const activeDeck = session ? allDecks.find((deck) => deck.id === session.deckId) : undefined;
  const renamingDeck = renamingDeckId ? allDecks.find((deck) => deck.id === renamingDeckId) : null;
  const todayReviewedWordIds = useMemo(
    () => new Set(dailyReviewedWords[getTodayKey()] ?? []),
    [dailyReviewedWords],
  );

  useEffect(() => {
    if (!toastMessage) return;
    const timer = window.setTimeout(() => setToastMessage(""), 1800);
    return () => window.clearTimeout(timer);
  }, [toastMessage]);

  useEffect(() => {
    let cancelled = false;

    getSupabaseSession()
      .then((session) => {
        if (!session?.user.id) {
          setAuthReady(false);
          setCheckingAuth(false);
          return null;
        }
        setAuthReady(true);
        setCurrentUserEmail(session.user.email ?? "");
        supabaseUserIdRef.current = session.user.id;
        setActiveStorageUserId(session.user.id);
        return loadSupabaseData(session.user.id).then((remoteData) => ({ remoteData, userId: session.user.id }));
      })
      .then((result) => {
        if (cancelled) return;
        if (!result) return;
        const { remoteData, userId } = result;
        return initializeAccountData(remoteData, userId);
      })
      .catch(() => {
        if (cancelled) return;
        setAuthReady(false);
        setToastMessage("登录状态检查失败");
      })
      .finally(() => {
        if (!cancelled) setCheckingAuth(false);
      });

    const { data: authListener } = supabase.auth.onAuthStateChange((_event, session) => {
      if (!session?.user.id) {
        supabaseUserIdRef.current = null;
        setActiveStorageUserId(null);
        setAuthReady(false);
        setCurrentUserEmail("");
        return;
      }
      supabaseUserIdRef.current = session.user.id;
      setActiveStorageUserId(session.user.id);
      setAuthReady(true);
      setCurrentUserEmail(session.user.email ?? "");
      loadSupabaseData(session.user.id)
        .then((remoteData) => {
          return initializeAccountData(remoteData, session.user.id);
        })
        .catch((error) => {
          console.error("Supabase hydration failed", error);
          setToastMessage("Supabase 数据读取失败");
        });
    });

    return () => {
      cancelled = true;
      if (supabaseSyncTimerRef.current) window.clearTimeout(supabaseSyncTimerRef.current);
      authListener.subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    setSpeechSupported(canUseSpeechSynthesis());
    if (typeof window === "undefined" || !("speechSynthesis" in window)) return;
    window.speechSynthesis.getVoices();

    function unlockSpeechSynthesis() {
      window.speechSynthesis.resume();
      window.speechSynthesis.getVoices();
    }

    window.addEventListener("pointerdown", unlockSpeechSynthesis, { once: true });
    window.addEventListener("keydown", unlockSpeechSynthesis, { once: true });
    return () => {
      window.removeEventListener("pointerdown", unlockSpeechSynthesis);
      window.removeEventListener("keydown", unlockSpeechSynthesis);
    };
  }, []);

  useEffect(() => {
    const speechWindow = window as SpeechErrorWindow;
    speechWindow.__wordloopShowSpeechError = (message) => {
      setToastMessage(message ?? "当前浏览器不支持系统语音播放，请用 Chrome 打开后重试");
    };

    return () => {
      delete speechWindow.__wordloopShowSpeechError;
    };
  }, []);

  useEffect(() => {
    if (!openDeckMenuId) return;

    function closeDeckMenuOnOutsidePointer(event: PointerEvent) {
      const target = event.target as HTMLElement | null;
      if (target?.closest(".deck-more-menu") || target?.closest(".more-button")) return;
      setOpenDeckMenuId(null);
    }

    window.addEventListener("pointerdown", closeDeckMenuOnOutsidePointer);
    return () => window.removeEventListener("pointerdown", closeDeckMenuOnOutsidePointer);
  }, [openDeckMenuId]);

  function withSettingsModal(page: ReactNode) {
    return (
      <>
        {page}
        {settingsOpen ? (
          <SettingsModal
            initialSettings={settings}
            onClose={() => setSettingsOpen(false)}
            onSave={(nextSettings) => {
              persistSettings(nextSettings);
              setSettingsOpen(false);
            }}
          />
        ) : null}
        {quickModalOpen ? (
          <QuickWordModal
            apiKey={settings.deepseekApiKey}
            decks={allDecks}
            onClose={() => setQuickModalOpen(false)}
            onCreateDeck={createEmptyUserDeck}
            onSave={saveQuickWords}
          />
        ) : null}
        {importModalOpen ? (
          <ImportDeckModal
            activeTab={importTab}
            apiKey={settings.deepseekApiKey}
            onClose={() => setImportModalOpen(false)}
            onCreateDeck={createUserDeck}
            onTabChange={setImportTab}
          />
        ) : null}
        {renamingDeck ? (
          <RenameDeckDialog
            deck={renamingDeck}
            existingNames={allDecks.filter((deck) => deck.id !== renamingDeck.id).map((deck) => deck.name)}
            onClose={() => setRenamingDeckId(null)}
            onSave={renameUserDeck}
          />
        ) : null}
        {expandingDeckId ? (
          <ExpandDeckDialog
            apiKey={settings.deepseekApiKey}
            deck={allDecks.find((deck) => deck.id === expandingDeckId) ?? null}
            onClose={() => setExpandingDeckId(null)}
            onExpand={expandDeckWithAi}
          />
        ) : null}
        {pendingDeckConfirm ? (
          <DeckConfirmDialog
            action={pendingDeckConfirm.action}
            deck={pendingDeckConfirm.deck}
            onClose={() => setPendingDeckConfirm(null)}
            onConfirm={() => confirmDeckAction(pendingDeckConfirm)}
          />
        ) : null}
        {importWordsDeckId ? (
          <ImportWordsModal
            apiKey={settings.deepseekApiKey}
            deck={allDecks.find((deck) => deck.id === importWordsDeckId) ?? null}
            onClose={() => setImportWordsDeckId(null)}
            onSave={importWordsToDeck}
          />
        ) : null}
        {toastMessage ? <AppToast message={toastMessage} /> : null}
      </>
    );
  }

  function queueSupabaseSync() {
    if (!supabaseUserIdRef.current) return;
    if (supabaseSyncTimerRef.current) window.clearTimeout(supabaseSyncTimerRef.current);
    supabaseSyncTimerRef.current = window.setTimeout(() => {
      const userId = supabaseUserIdRef.current;
      if (!userId) return;
      syncLocalDataToSupabase(userId).catch((error) => {
        console.error("Supabase sync failed", error);
        setToastMessage("Supabase 同步失败");
      });
    }, 500);
  }

  function syncSupabaseNow() {
    const userId = supabaseUserIdRef.current;
    if (!userId) return;
    if (supabaseSyncTimerRef.current) {
      window.clearTimeout(supabaseSyncTimerRef.current);
      supabaseSyncTimerRef.current = null;
    }
    syncLocalDataToSupabase(userId).catch((error) => {
      console.error("Supabase sync failed", error);
      setToastMessage("Supabase 同步失败");
    });
  }

  function hasHydrationData(data: SupabaseHydrationData | null) {
    if (!data) return false;
    return [
      data.userDecks,
      data.deckOverrides,
      data.progress,
      data.mastery,
      data.reviewCycles,
      data.reviewStats,
      data.reviewSchedule,
      data.dailyReviewedWords,
    ].some(hasMeaningfulStoredValue);
  }

  function hasLocalHydrationData() {
    return [
      loadUserDecks(),
      loadDeckOverrides(),
      loadProgress(),
      loadMastery(),
      loadReviewCycles(),
      loadReviewStats(),
      loadReviewSchedule(),
      loadDailyReviewedWords(),
    ].some(hasMeaningfulStoredValue);
  }

  async function initializeAccountData(remoteData: SupabaseHydrationData | null, userId: string) {
    if (hasHydrationData(remoteData)) {
      hydrateFromSupabase(remoteData as SupabaseHydrationData);
      setToastMessage("已从 Supabase 加载数据");
      return;
    }

    copyLegacyStorageToCurrentUser([
      userDecksStorageKey,
      deckOverridesStorageKey,
      progressStorageKey,
      masteryStorageKey,
      reviewCyclesStorageKey,
      reviewStatsStorageKey,
      reviewScheduleStorageKey,
      dailyReviewedWordsStorageKey,
      settingsStorageKey,
    ]);

    if (hasLocalHydrationData()) {
      hydrateFromLocalStorage();
      await syncLocalDataToSupabase(userId);
      setToastMessage("已完成首次本地数据导入");
      return;
    }

    hydrateFromSupabase(
      remoteData ?? {
        userDecks: [],
        progress: createEmptyProgress(),
        mastery: createEmptyMastery(),
        reviewCycles: createEmptyReviewCycles(),
        reviewStats: {},
        dailyReviewedWords: {},
        deckOverrides: {},
        reviewSchedule: {},
        settings: {},
      },
    );
    await syncLocalDataToSupabase(userId);
    setToastMessage("已初始化 Supabase 数据");
  }

  function hydrateFromSupabase(remoteData: SupabaseHydrationData) {
    const nextSettings = { ...loadSettings(), ...remoteData.settings };
    setUserDecks(remoteData.userDecks);
    saveUserDecks(remoteData.userDecks);
    setDeckOverrides(remoteData.deckOverrides);
    saveDeckOverrides(remoteData.deckOverrides);
    setProgress(remoteData.progress);
    saveProgress(remoteData.progress);
    setMastery(remoteData.mastery);
    saveMastery(remoteData.mastery);
    setReviewCycles(remoteData.reviewCycles);
    saveReviewCycles(remoteData.reviewCycles);
    setReviewStats(remoteData.reviewStats);
    saveReviewStats(remoteData.reviewStats);
    setReviewSchedule(remoteData.reviewSchedule);
    saveReviewSchedule(remoteData.reviewSchedule);
    setDailyReviewedWords(remoteData.dailyReviewedWords);
    saveDailyReviewedWords(remoteData.dailyReviewedWords);
    setSettings(nextSettings);
    saveSettings(nextSettings);
  }

  function hydrateFromLocalStorage() {
    const nextUserDecks = loadUserDecks();
    const nextDeckOverrides = loadDeckOverrides();
    const nextProgress = loadProgress();
    const nextMastery = loadMastery();
    const nextReviewCycles = loadReviewCycles();
    const nextReviewStats = loadReviewStats();
    const nextReviewSchedule = loadReviewSchedule();
    const nextDailyReviewedWords = loadDailyReviewedWords();
    const nextSettings = loadSettings();
    setUserDecks(nextUserDecks);
    setDeckOverrides(nextDeckOverrides);
    setProgress(nextProgress);
    setMastery(nextMastery);
    setReviewCycles(nextReviewCycles);
    setReviewStats(nextReviewStats);
    setReviewSchedule(nextReviewSchedule);
    setDailyReviewedWords(nextDailyReviewedWords);
    setSettings(nextSettings);
  }

  async function requestPasswordAuth(email: string, password: string, mode: AuthMode) {
    if (mode === "signUp") {
      const session = await signUpWithPassword(email, password);
      setToastMessage(session ? "注册成功，已登录" : "注册成功，请先确认邮箱后再登录");
      return;
    }
    await signInWithPassword(email, password);
    setToastMessage("已登录");
  }

  async function requestAnonymousLogin() {
    await signInAnonymously();
    setToastMessage("已匿名登录");
  }

  async function signOut() {
    await signOutSupabase();
    supabaseUserIdRef.current = null;
    setActiveStorageUserId(null);
    setAuthReady(false);
    setCurrentUserEmail("");
    setToastMessage("已退出登录");
  }

  function persistProgress(nextProgress: StoredProgress) {
    setProgress(nextProgress);
    saveProgress(nextProgress);
    queueSupabaseSync();
  }

  function persistMastery(nextMastery: StoredMastery) {
    setMastery(nextMastery);
    saveMastery(nextMastery);
    queueSupabaseSync();
  }

  function persistReviewCycles(nextReviewCycles: StoredReviewCycles) {
    setReviewCycles(nextReviewCycles);
    saveReviewCycles(nextReviewCycles);
    queueSupabaseSync();
  }

  function persistReviewStats(nextReviewStats: StoredReviewStats) {
    setReviewStats(nextReviewStats);
    saveReviewStats(nextReviewStats);
    queueSupabaseSync();
  }

  function persistDailyReviewedWords(nextDailyReviewedWords: StoredDailyReviewedWords) {
    setDailyReviewedWords(nextDailyReviewedWords);
    saveDailyReviewedWords(nextDailyReviewedWords);
    queueSupabaseSync();
  }

  function persistReviewSchedule(nextReviewSchedule: StoredReviewSchedule) {
    setReviewSchedule(nextReviewSchedule);
    saveReviewSchedule(nextReviewSchedule);
    queueSupabaseSync();
  }

  function persistDeckOverrides(nextDeckOverrides: StoredDeckOverrides) {
    setDeckOverrides(nextDeckOverrides);
    saveDeckOverrides(nextDeckOverrides);
    queueSupabaseSync();
  }

  function persistUserDecks(nextUserDecks: VocabularyDeck[]) {
    setUserDecks(nextUserDecks);
    saveUserDecks(nextUserDecks);
    queueSupabaseSync();
  }

  function createUserDeck(name: string, words: Array<string | ImportedWordInput>, description = "本地创建词库", aiGenerated = false) {
    const deckName = name.trim() || "自定义词库";
    const existing = userDecks.find((deck) => deck.name === deckName);
    const deckId = existing?.id ?? `custom-${slugifyDeckName(deckName) || Date.now()}`;
    const existingWords = new Set((existing?.words ?? []).map((word) => word.word.trim().toLowerCase()));
    const uniqueWords = words.filter((word) => {
      const normalized = getImportedWordText(word).toLowerCase();
      if (!normalized || existingWords.has(normalized)) return false;
      existingWords.add(normalized);
      return true;
    });
    const nextWords = uniqueWords.map((word, index) =>
      createLocalWord(deckId, word, (existing?.words.length ?? 0) + index + 1),
    );
    if (!nextWords.length) {
      setToastMessage("没有可添加的新单词");
      return 0;
    }

    const nextDeck: VocabularyDeck = existing
      ? { ...existing, aiGenerated: existing.aiGenerated || aiGenerated, words: [...existing.words, ...nextWords] }
      : {
          id: deckId,
          name: deckName,
          description,
          words: nextWords,
          aiGenerated,
          custom: true,
        };
    persistUserDecks(existing ? userDecks.map((deck) => (deck.id === existing.id ? nextDeck : deck)) : [...userDecks, nextDeck]);
    return nextWords.length;
  }

  function createEmptyUserDeck(name: string) {
    const deckName = name.trim();
    if (!deckName) return "";
    const existingDeck = allDecks.find((deck) => deck.name === deckName);
    if (existingDeck) return existingDeck.name;
    const deckId = `custom-${slugifyDeckName(deckName) || Date.now()}`;
    const nextDeck: VocabularyDeck = {
      id: deckId,
      name: deckName,
      description: "手动创建的空词库。",
      words: [],
      custom: true,
    };
    persistUserDecks([...userDecks, nextDeck]);
    return deckName;
  }

  function saveQuickWords(words: Array<string | ImportedWordInput>, deckName: string) {
    const addedCount = createUserDeck(deckName || "速记单词", words, "快速记录的临时词库。");
    if (addedCount) setQuickModalOpen(false);
  }

  function expandDeckWithAi(deck: VocabularyDeck, words: ImportedWordInput[]) {
    const addedCount = createUserDeck(deck.name, words, "通过 AI 扩充的词库。", true);
    if (!addedCount) return;
    setExpandingDeckId(null);
    setOpenDeckMenuId(null);
  }

  function importWordsToDeck(deck: VocabularyDeck, words: Array<string | ImportedWordInput>) {
    const targetName = deck.custom ? deck.name : `${deck.name}导入词库`;
    const addedCount = createUserDeck(targetName, words, `导入至「${deck.name}」的补充词库。`);
    if (!addedCount) return;
    setImportWordsDeckId(null);
    setOpenDeckMenuId(null);
  }

  function renameUserDeck(deckId: string, nextName: string) {
    const deckName = nextName.trim();
    if (!deckName) return;
    const targetDeck = userDecks.find((deck) => deck.id === deckId);
    if (targetDeck) {
      persistUserDecks(userDecks.map((deck) => (deck.id === deckId ? { ...deck, name: deckName } : deck)));
    } else {
      const builtinDeck = builtinDecks.find((deck) => deck.id === deckId);
      if (!builtinDeck) return;
      persistDeckOverrides({
        ...deckOverrides,
        [deckId]: {
          ...deckOverrides[deckId],
          name: deckName,
        },
      });
    }
    setRenamingDeckId(null);
    setOpenDeckMenuId(null);
  }

  function resetDeck(deckId: string) {
    const nextProgress = { ...progress };
    const nextMastery = { ...mastery };
    const nextReviewCycles = { ...reviewCycles };
    const nextReviewSchedule = { ...reviewSchedule };
    nextProgress[deckId] = {};
    nextMastery[deckId] = {};
    nextReviewCycles[deckId] = {};
    nextReviewSchedule[deckId] = {};
    persistProgress(nextProgress);
    persistMastery(nextMastery);
    persistReviewCycles(nextReviewCycles);
    persistReviewSchedule(nextReviewSchedule);
    const nextDailyReviewedWords = Object.fromEntries(
      Object.entries(dailyReviewedWords).map(([date, wordIds]) => [
        date,
        wordIds.filter((wordId) => !allDecks.find((deck) => deck.id === deckId)?.words.some((word) => word.id === wordId)),
      ]),
    );
    persistDailyReviewedWords(nextDailyReviewedWords);
    setOpenDeckMenuId(null);
    syncSupabaseNow();
  }

  function deleteDeck(deckId: string) {
    const targetDeck = allDecks.find((deck) => deck.id === deckId);
    if (!targetDeck) return;
    if (targetDeck.custom) {
      persistUserDecks(userDecks.filter((deck) => deck.id !== deckId));
    } else {
      persistDeckOverrides({
        ...deckOverrides,
        [deckId]: {
          ...deckOverrides[deckId],
          deleted: true,
        },
      });
    }

    const nextProgress = { ...progress };
    const nextMastery = { ...mastery };
    const nextReviewCycles = { ...reviewCycles };
    const nextReviewSchedule = { ...reviewSchedule };
    delete nextProgress[deckId];
    delete nextMastery[deckId];
    delete nextReviewCycles[deckId];
    delete nextReviewSchedule[deckId];
    persistProgress(nextProgress);
    persistMastery(nextMastery);
    persistReviewCycles(nextReviewCycles);
    persistReviewSchedule(nextReviewSchedule);
    const deletedWordIds = new Set(targetDeck.words.map((word) => word.id));
    const nextDailyReviewedWords = Object.fromEntries(
      Object.entries(dailyReviewedWords).map(([date, wordIds]) => [
        date,
        wordIds.filter((wordId) => !deletedWordIds.has(wordId)),
      ]),
    );
    persistDailyReviewedWords(nextDailyReviewedWords);
    setOpenDeckMenuId(null);
    syncSupabaseNow();
  }

  function confirmDeckAction(confirm: PendingDeckConfirm) {
    if (confirm.action === "reset") resetDeck(confirm.deck.id);
    if (confirm.action === "delete") deleteDeck(confirm.deck.id);
    setPendingDeckConfirm(null);
  }

  function exportDeck(deck: VocabularyDeck) {
    downloadDeckAsExcel(deck);
    setOpenDeckMenuId(null);
    setToastMessage("开始下载");
  }

  function completeReviewWords(words: VocabularyWord[], status: MarkStatus) {
    if (!words.length) return;
    const nextMastery: StoredMastery = { ...mastery };
    const nextReviewCycles: StoredReviewCycles = { ...reviewCycles };
    const nextReviewSchedule: StoredReviewSchedule = { ...reviewSchedule };
    const todayKey = getTodayKey();
    const nextReviewStats = {
      ...reviewStats,
      [todayKey]: (reviewStats[todayKey] ?? 0) + words.length,
    };
    const nextTodayReviewedWords = Array.from(
      new Set([...(dailyReviewedWords[todayKey] ?? []), ...words.map((word) => word.id)]),
    );
    const nextDailyReviewedWords = {
      ...dailyReviewedWords,
      [todayKey]: nextTodayReviewedWords,
    };
    let changed = false;
    let reviewCyclesChanged = false;

    words.forEach((word) => {
      const deckId = getDeckIdForWord(word.id, allDecks);
      if (!deckId) return;
      const currentMemory = getWordMemory(deckId, word.id, progress, nextReviewCycles, nextReviewSchedule);
      const nextMemory = scheduleAfterReview(currentMemory, status);
      const currentCycles = nextReviewCycles[deckId]?.[word.id] ?? 0;
      const nextCycles = currentCycles + 1;
      nextReviewCycles[deckId] = { ...nextReviewCycles[deckId], [word.id]: nextCycles };
      nextReviewSchedule[deckId] = { ...nextReviewSchedule[deckId], [word.id]: nextMemory };
      reviewCyclesChanged = true;
      if (isMemoryMastered(nextMemory)) {
        nextMastery[deckId] = { ...nextMastery[deckId], [word.id]: true };
        changed = true;
      } else if (nextMastery[deckId]?.[word.id]) {
        const nextDeckMastery = { ...nextMastery[deckId] };
        delete nextDeckMastery[word.id];
        nextMastery[deckId] = nextDeckMastery;
        changed = true;
      }
    });

    if (reviewCyclesChanged) persistReviewCycles(nextReviewCycles);
    if (reviewCyclesChanged) persistReviewSchedule(nextReviewSchedule);
    if (changed) persistMastery(nextMastery);
    persistDailyReviewedWords(nextDailyReviewedWords);
    persistReviewStats(nextReviewStats);
  }

  function startStudy(deckId: string) {
    const deck = allDecks.find((item) => item.id === deckId);
    if (!deck) return;
    setSession(createSession(deck, progress, settings.dailyStudyCount));
    setExpandedWordIds(new Set());
    setQuizNotice("");
    window.location.hash = `/study/${deckId}`;
    setView("study");
  }

  function backHome() {
    window.location.hash = "";
    setView("home");
    setQuizNotice("");
  }

  function startReview() {
    window.location.hash = "/review";
    setView("review");
  }

  function markWord(wordId: string, status: MarkStatus) {
    if (!session) return;
    const nextSession = {
      ...session,
      statuses: { ...session.statuses, [wordId]: status },
    };
    setSession(nextSession);
  }

  function undoWord(wordId: string) {
    if (!session) return;
    const nextStatuses = { ...session.statuses };
    delete nextStatuses[wordId];

    setSession({ ...session, statuses: nextStatuses });
  }

  function completeCurrentLearningSession(result: TranslationEvaluationResult) {
    if (!session || !activeDeck) return;
    const learnedStatuses = Object.entries(session.statuses);
    if (!learnedStatuses.length) return;
    const resultByWord = new Map(result.reviewRows.map((row) => [row.word.trim().toLowerCase(), row.statusType]));
    const nextReviewSchedule: StoredReviewSchedule = { ...reviewSchedule };
    const deckSchedule = { ...(nextReviewSchedule[session.deckId] ?? {}) };
    learnedStatuses.forEach(([wordId, status]) => {
      const word = activeDeck.words.find((item) => item.id === wordId);
      if (!word) return;
      deckSchedule[wordId] = scheduleAfterStudyAndQuiz(
        status,
        resultByWord.get(word.word.trim().toLowerCase()),
        deckSchedule[wordId],
      );
    });
    nextReviewSchedule[session.deckId] = deckSchedule;
    persistProgress({
      ...progress,
      [session.deckId]: {
        ...progress[session.deckId],
        ...Object.fromEntries(learnedStatuses),
      },
    });
    persistReviewSchedule(nextReviewSchedule);
  }

  function addTenWords() {
    if (!session || !activeDeck) return;
    const current = new Set(session.wordIds);
    const deckProgress = progress[session.deckId] ?? {};
    const nextWords = activeDeck.words.filter((word) => !current.has(word.id) && !deckProgress[word.id]).slice(0, 10);
    if (!nextWords.length) return;
    setSession({
      ...session,
      wordIds: [...session.wordIds, ...nextWords.map((word) => word.id)],
    });
  }

  function getQuizWords() {
    if (!session || !activeDeck) return [];
    const sessionWords = session.wordIds
      .map((id) => activeDeck.words.find((word) => word.id === id))
      .filter(Boolean) as VocabularyWord[];
    const markedWords = sessionWords.filter((word) => session.statuses[word.id]);
    return markedWords.length ? markedWords : sessionWords;
  }

  function createMockTranslationPrompt(mode: TranslationMode, targetWords: VocabularyWord[]): TranslationPromptData {
    const words = targetWords.map((word) => word.word);
    if (mode === "zhToEn") {
      return {
        source: `请把下面这段话翻译成英文：团队正在复盘一次产品设计会议，大家希望用更务实的方式处理短期反馈，同时保持长期目标的韧性。请尽量自然地使用这些表达：${words.join("、")}。`,
        highlightedWords: words,
      };
    }

    return {
      source: `During the design review, the team tried to stay pragmatic while discussing a complex user flow. They wanted every decision to be resilient enough for future changes, but still clear for today's release. Several details felt ephemeral, so the designer asked everyone to focus on the vocabulary and the real user context.`,
      highlightedWords: words,
    };
  }

  function createMockTranslationResult(source: string, userTranslation: string, targetWords: string[]): TranslationEvaluationResult {
    return {
      score: 82,
      source,
      critique: "本地 mock 测评：整体表达完整，核心语义基本覆盖。部分词可以再贴近上下文，而不是逐词直译。",
      userTranslation,
      reference: "这是本地生成的参考译文，用于在未配置 API Key 时测试测验流程。",
      reviewRows: targetWords.map((word, index) => ({
        word,
        meaning: index % 3 === 0 ? "语义准确" : index % 3 === 1 ? "基本理解，语境略偏" : "建议结合上下文复习",
        status: index % 3 === 0 ? "正确" : index % 3 === 1 ? "基本理解" : "待复习",
        statusType: index % 3 === 0 ? "correct" : index % 3 === 1 ? "partial" : "review",
        reason: index % 3 === 0 ? "译文覆盖了核心意思。" : index % 3 === 1 ? "大意接近，但搭配或语境可以更自然。" : "表达略生硬，建议回到词卡例句再看一遍。",
      })),
    };
  }

  async function generatePromptForMode(mode: TranslationMode) {
    const targetWords = getQuizWords();
    const targetWordTexts = targetWords.map((word) => word.word);
    if (!targetWords.length) {
      setTranslationError("当前学习会话没有可用于生成测验的单词。");
      return null;
    }
    setTranslationTargetWords(targetWordTexts);

    if (!settings.deepseekApiKey.trim()) {
      const prompt = createMockTranslationPrompt(mode, targetWords);
      setTranslationPrompt(prompt);
      setTranslationError("");
      return prompt;
    }

    setTranslationStatus("generating");
    setTranslationError("");
    try {
      const prompt = await generateTranslationPromptWithDeepSeek({
        apiKey: settings.deepseekApiKey.trim(),
        mode,
        targetWords,
      });
      setTranslationPrompt(prompt);
      return prompt;
    } catch (error) {
      setTranslationPrompt(null);
      setTranslationError(error instanceof Error ? error.message : "DeepSeek 生成测验段落失败。");
      return null;
    } finally {
      setTranslationStatus("idle");
    }
  }

  async function enterQuiz() {
    if (!session) return;
    const markedCount = Object.keys(session.statuses).length;
    if (!markedCount) {
      setQuizNotice("至少标记 1 个单词后才可以进入测验。");
      return;
    }
    setTranslationInput("");
    setTranslationResult(null);
    setTranslationPrompt(null);
    setTranslationTargetWords([]);
    setTranslationError("");
    window.location.hash = "/ai-translate";
    setView("translate");
    const prompt = await generatePromptForMode(translationMode);
    if (!prompt) {
      return;
    }
  }

  function persistSettings(nextSettings: AppSettings) {
    setSettings(nextSettings);
    saveSettings(nextSettings);
    queueSupabaseSync();
  }

  async function submitTranslation() {
    if (!translationInput.trim()) return;

    setTranslationStatus("evaluating");
    setTranslationError("");
    try {
      if (!translationPrompt) throw new Error("请先生成测验段落。");
      const evaluationTargetWords = translationTargetWords.length
        ? translationTargetWords
        : translationPrompt.highlightedWords;
      if (!settings.deepseekApiKey.trim()) {
        const result = createMockTranslationResult(
          translationPrompt.source,
          translationInput.trim(),
          evaluationTargetWords,
        );
        completeCurrentLearningSession(result);
        setTranslationResult(result);
        window.location.hash = "/ai-translate/result";
        setView("translateResult");
        return;
      }

      const result = await evaluateTranslationWithDeepSeek({
        apiKey: settings.deepseekApiKey.trim(),
        mode: translationMode,
        source: translationPrompt.source,
        userTranslation: translationInput.trim(),
        targetWords: evaluationTargetWords,
      });
      completeCurrentLearningSession(result);
      setTranslationResult(result);
      window.location.hash = "/ai-translate/result";
      setView("translateResult");
    } catch (error) {
      setTranslationError(error instanceof Error ? error.message : "DeepSeek 测评失败。");
    } finally {
      setTranslationStatus("idle");
    }
  }

  function retryTranslation() {
    setTranslationInput("");
    setTranslationError("");
    setTranslationResult(null);
    window.location.hash = "/ai-translate";
    setView("translate");
  }

  async function changeTranslationMode(nextMode: TranslationMode) {
    setTranslationMode(nextMode);
    setTranslationInput("");
    setTranslationResult(null);
    if (view === "translate") await generatePromptForMode(nextMode);
  }

  function completeTranslation() {
    window.location.hash = "";
    setView("home");
  }

  async function ensureWordDetails(deckId: string, word: VocabularyWord) {
    const detailKey = getWordDetailKey(deckId, word.id);
    if (wordDetailsById[detailKey]?.data || wordDetailsById[detailKey]?.loading) return;

    const userId = supabaseUserIdRef.current;
    if (!userId) {
      setWordDetailsById((current) => ({
        ...current,
        [detailKey]: { error: "请先登录后再查看详情。" },
      }));
      return;
    }

    const apiKey = settings.deepseekApiKey.trim();
    if (!apiKey) {
      setWordDetailsById((current) => ({
        ...current,
        [detailKey]: { error: "请先在设置中配置 DeepSeek API Key。" },
      }));
      return;
    }

    setWordDetailsById((current) => ({ ...current, [detailKey]: { loading: true } }));
    try {
      const cachedDetails = await loadCachedWordDetails(userId, deckId, word.id);
      if (cachedDetails) {
        setWordDetailsById((current) => ({ ...current, [detailKey]: { data: cachedDetails } }));
        return;
      }
      const generatedDetails = await generateWordDetailsWithDeepSeek(apiKey, word);
      await saveCachedWordDetails(userId, deckId, word.id, generatedDetails);
      setWordDetailsById((current) => ({ ...current, [detailKey]: { data: generatedDetails } }));
    } catch (error) {
      console.error("Word details generation failed", error);
      setWordDetailsById((current) => ({
        ...current,
        [detailKey]: { error: error instanceof Error ? error.message : "详情生成失败，请稍后重试。" },
      }));
    }
  }

  function toggleExpanded(word: VocabularyWord) {
    if (!session) return;
    const shouldLoad = !expandedWordIds.has(word.id);
    setExpandedWordIds((current) => {
      const next = new Set(current);
      if (next.has(word.id)) next.delete(word.id);
      else next.add(word.id);
      return next;
    });
    if (shouldLoad) void ensureWordDetails(session.deckId, word);
  }

  if (checkingAuth) {
    return <AuthStatusPage message="正在检查登录状态" />;
  }

  if (!authReady) {
    return (
      <>
        <EmailLoginPage onAnonymousLogin={requestAnonymousLogin} onSubmit={requestPasswordAuth} />
        {toastMessage ? <AppToast message={toastMessage} /> : null}
      </>
    );
  }

  if (view === "study" && session && activeDeck) {
    return withSettingsModal(
      <StudyPage
        deck={activeDeck}
        expandedWordIds={expandedWordIds}
        onAddTenWords={addTenWords}
        onBack={backHome}
        onEnterQuiz={enterQuiz}
        onMarkWord={markWord}
        onToggleExpanded={toggleExpanded}
        onUndoWord={undoWord}
        quizNotice={quizNotice}
        session={session}
        speechSupported={speechSupported}
        wordDetailsById={wordDetailsById}
      />,
    );
  }

  if (view === "translate") {
    return withSettingsModal(
      <TranslationTestPage
        onBack={
          session
            ? () => {
                window.location.hash = `/study/${session.deckId}`;
                setView("study");
              }
            : backHome
        }
        onSubmit={submitTranslation}
        onModeChange={changeTranslationMode}
        onOpenSettings={() => setSettingsOpen(true)}
        onTranslationInputChange={setTranslationInput}
        prompt={translationPrompt}
        settingsConfigured={Boolean(settings.deepseekApiKey.trim())}
        status={translationStatus}
        translationError={translationError}
        translationInput={translationInput}
        translationMode={translationMode}
      />,
    );
  }

  if (view === "translateResult") {
    return withSettingsModal(
      <TranslationResultPage
        onBack={
          session
            ? () => {
                window.location.hash = `/study/${session.deckId}`;
                setView("study");
              }
            : backHome
        }
        onComplete={completeTranslation}
        onRetry={retryTranslation}
        result={translationResult}
      />,
    );
  }

  if (view === "review") {
    return withSettingsModal(
      <ReviewPage
        decks={allDecks}
        deckFilter={reviewDeckFilter}
        excludedWordIds={todayReviewedWordIds}
        getWordDetailState={(word) => {
          const deckId = getDeckIdForWord(word.id, allDecks);
          return deckId ? wordDetailsById[getWordDetailKey(deckId, word.id)] : undefined;
        }}
        mastery={mastery}
        onBack={backHome}
        onDeckFilterChange={setReviewDeckFilter}
        onEnsureWordDetails={(word) => {
          const deckId = getDeckIdForWord(word.id, allDecks);
          if (deckId) void ensureWordDetails(deckId, word);
        }}
        onReviewWordComplete={(word, status) => completeReviewWords([word], status)}
        progress={progress}
        reviewCycles={reviewCycles}
        reviewSchedule={reviewSchedule}
        reviewLimit={settings.dailyReviewGoal}
        speechSupported={speechSupported}
      />,
    );
  }

  return withSettingsModal(
      <HomePage
        decks={allDecks}
        onOpenSettings={() => setSettingsOpen(true)}
        onOpenImport={() => {
          setImportTab("ai");
          setImportModalOpen(true);
        }}
        onOpenQuick={() => setQuickModalOpen(true)}
        onConfirmDeckAction={(deck, action) => {
          setPendingDeckConfirm({ action, deck });
          setOpenDeckMenuId(null);
        }}
        onExportDeck={exportDeck}
        onExpandDeck={(deckId) => {
          setExpandingDeckId(deckId);
          setOpenDeckMenuId(null);
        }}
        onImportWords={(deckId) => {
          setImportWordsDeckId(deckId);
          setOpenDeckMenuId(null);
        }}
        onStartReview={startReview}
        onStartStudy={startStudy}
        onRenameDeck={(deckId) => {
          setRenamingDeckId(deckId);
          setOpenDeckMenuId(null);
        }}
        onToggleDeckMenu={(deckId) => setOpenDeckMenuId((current) => (current === deckId ? null : deckId))}
        openDeckMenuId={openDeckMenuId}
        mastery={mastery}
        progress={progress}
        reviewCycles={reviewCycles}
        reviewSchedule={reviewSchedule}
        reviewStats={reviewStats}
        settings={settings}
        userEmail={currentUserEmail}
        onSignOut={signOut}
      />,
  );
}

function HomePage({
  decks,
  mastery,
  onOpenImport,
  onOpenSettings,
  onOpenQuick,
  onConfirmDeckAction,
  onExportDeck,
  onExpandDeck,
  onImportWords,
  onRenameDeck,
  onStartReview,
  onStartStudy,
  onToggleDeckMenu,
  onSignOut,
  openDeckMenuId,
  progress,
  reviewCycles,
  reviewSchedule,
  reviewStats,
  settings,
  userEmail,
}: {
  decks: VocabularyDeck[];
  mastery: StoredMastery;
  onOpenImport: () => void;
  onOpenSettings: () => void;
  onOpenQuick: () => void;
  onConfirmDeckAction: (deck: VocabularyDeck, action: DeckConfirmAction) => void;
  onExportDeck: (deck: VocabularyDeck) => void;
  onExpandDeck: (deckId: string) => void;
  onImportWords: (deckId: string) => void;
  onRenameDeck: (deckId: string) => void;
  onStartReview: () => void;
  onStartStudy: (deckId: string) => void;
  onToggleDeckMenu: (deckId: string) => void;
  onSignOut: () => void;
  openDeckMenuId: string | null;
  progress: StoredProgress;
  reviewCycles: StoredReviewCycles;
  reviewSchedule: StoredReviewSchedule;
  reviewStats: StoredReviewStats;
  settings: AppSettings;
  userEmail: string;
}) {
  const homeReview = getHomeReviewData(progress, mastery, reviewCycles, reviewSchedule, decks, settings.dailyReviewGoal, reviewStats);
  const learningStats = getLearningRecordStats(progress, mastery, reviewCycles, reviewSchedule, decks);
  const sortedDecks = useMemo(
    () =>
      decks
        .map((deck, index) => ({
          deck,
          index,
          complete: isDeckLearningComplete(deck, progress[deck.id] ?? {}),
        }))
        .sort((a, b) => Number(a.complete) - Number(b.complete) || a.index - b.index),
    [decks, progress],
  );

  return (
    <main className="home-page">
      <Header onOpenSettings={onOpenSettings} onSignOut={onSignOut} userEmail={userEmail} />
      <div className="home-main">
        <section className="learning-section" aria-labelledby="learning-title">
          <div className="section-header">
            <h1 id="learning-title">学习</h1>
            <div className="section-actions">
              <AppButton tone="outline" icon={<Zap size={18} />} onClick={onOpenQuick}>
                快速记单词
              </AppButton>
              <AppButton tone="primary" icon={<Import size={24} />} onClick={onOpenImport}>
                导入新词库
              </AppButton>
            </div>
          </div>
          <div className="deck-grid">
            {sortedDecks.map(({ complete, deck }) => (
              <DeckCard
                key={deck.id}
                deck={deck}
                learningComplete={complete}
                isMenuOpen={openDeckMenuId === deck.id}
                masteredWords={mastery[deck.id] ?? {}}
                onConfirmAction={onConfirmDeckAction}
                onExport={onExportDeck}
                onExpand={onExpandDeck}
                onImportWords={onImportWords}
                onRename={onRenameDeck}
                onToggleMenu={onToggleDeckMenu}
                onStartStudy={onStartStudy}
                progress={progress[deck.id] ?? {}}
                reviewCycles={reviewCycles[deck.id] ?? {}}
                reviewSchedule={reviewSchedule[deck.id] ?? {}}
              />
            ))}
          </div>
        </section>

        <aside className="sidebar" aria-label="复习与学习记录">
          <section className="review-section" aria-labelledby="review-title">
            <h2 id="review-title">复习</h2>
            <ReviewCard data={homeReview} onStartReview={onStartReview} />
          </section>
          <LearningRecordCard stats={learningStats} />
        </aside>
      </div>
    </main>
  );
}

function isDeckLearningComplete(deck: VocabularyDeck, progress: Record<string, MarkStatus>) {
  return deck.words.length > 0 && deck.words.every((word) => Boolean(progress[word.id]));
}

function getDeckSourceLabel(deck: VocabularyDeck) {
  if (!deck.custom) return "内置";
  if (deck.aiGenerated) return "AI 生成";
  return "本地上传";
}

function createReviewQueue(
  decks: VocabularyDeck[],
  deckFilter: ReviewDeckFilter,
  reviewLimit: number,
  progress: StoredProgress,
  mastery: StoredMastery,
  reviewCycles: StoredReviewCycles,
  reviewSchedule: StoredReviewSchedule,
  excludedWordIds: Set<string> = new Set(),
) {
  const sourceDecks =
    deckFilter === "all" ? decks : decks.filter((deck) => deck.id === deckFilter);
  return sourceDecks
    .flatMap((deck) =>
      deck.words
        .filter((word) => {
          if (!progress[deck.id]?.[word.id] || excludedWordIds.has(word.id)) return false;
          const memory = getWordMemory(deck.id, word.id, progress, reviewCycles, reviewSchedule);
          return !isMemoryMastered(memory) && isReviewDue(memory);
        })
        .map((word) => {
          const memory = getWordMemory(deck.id, word.id, progress, reviewCycles, reviewSchedule);
          const overdueHours = Math.max(0, (Date.now() - new Date(memory.dueAt).getTime()) / 36e5);
          const priority = overdueHours * 2 + (masteryReviewStageTarget - memory.stage) * 3 + memory.lapses * 2 + (1.2 - memory.ease);
          return { priority, word };
        }),
    )
    .sort((a, b) => b.priority - a.priority)
    .map(({ word }) => word)
    .slice(0, reviewLimit);
}

function ReviewPage({
  decks,
  deckFilter,
  excludedWordIds,
  getWordDetailState,
  mastery,
  onBack,
  onDeckFilterChange,
  onEnsureWordDetails,
  onReviewWordComplete,
  progress,
  reviewCycles,
  reviewSchedule,
  reviewLimit = reviewSessionLimit,
  speechSupported,
}: {
  decks: VocabularyDeck[];
  deckFilter: ReviewDeckFilter;
  excludedWordIds: Set<string>;
  getWordDetailState: (word: VocabularyWord) => WordDetailsState | undefined;
  mastery: StoredMastery;
  onBack: () => void;
  onDeckFilterChange: (deckFilter: ReviewDeckFilter) => void;
  onEnsureWordDetails: (word: VocabularyWord) => void;
  onReviewWordComplete: (word: VocabularyWord, status: MarkStatus) => void;
  progress: StoredProgress;
  reviewCycles: StoredReviewCycles;
  reviewSchedule: StoredReviewSchedule;
  reviewLimit?: number;
  speechSupported: boolean;
}) {
  const [queue, setQueue] = useState<VocabularyWord[]>(() =>
    createReviewQueue(decks, deckFilter, reviewLimit, progress, mastery, reviewCycles, reviewSchedule, excludedWordIds),
  );
  const [reviewIndex, setReviewIndex] = useState(0);
  const [flipped, setFlipped] = useState(false);
  const [detailsExpanded, setDetailsExpanded] = useState(false);
  const [exiting, setExiting] = useState(false);
  const [reviewStatuses, setReviewStatuses] = useState<MarkStatus[]>([]);
  const [reviewedWordIds, setReviewedWordIds] = useState<Set<string>>(() => new Set());
  const reviewedWordIdsRef = useRef<Set<string>>(new Set());
  const reviewStartedAtRef = useRef(Date.now());
  const activeWord = queue[reviewIndex];
  const completed = reviewIndex >= queue.length;
  const reviewedCount = Math.min(reviewIndex, queue.length);
  const quickRecallCount = reviewStatuses.filter((status) => status === "known").length;
  const quickRecallPercent = reviewStatuses.length ? Math.round((quickRecallCount / reviewStatuses.length) * 100) : 0;
  const focusMinutes = Math.max(1, Math.round((Date.now() - reviewStartedAtRef.current) / 60000));

  useEffect(() => {
    if (deckFilter !== "all" && !decks.some((deck) => deck.id === deckFilter)) {
      onDeckFilterChange("all");
      return;
    }
    setQueue(createReviewQueue(decks, deckFilter, reviewLimit, progress, mastery, reviewCycles, reviewSchedule, excludedWordIds));
    setReviewIndex(0);
    setFlipped(false);
    setDetailsExpanded(false);
    setExiting(false);
    setReviewStatuses([]);
    const emptyReviewedIds = new Set<string>();
    reviewedWordIdsRef.current = emptyReviewedIds;
    setReviewedWordIds(emptyReviewedIds);
    reviewStartedAtRef.current = Date.now();
  }, [deckFilter, reviewLimit, decks, onDeckFilterChange]);

  function changeDeckFilter(nextFilter: ReviewDeckFilter) {
    onDeckFilterChange(nextFilter);
  }

  function flipCard() {
    if (!activeWord || completed || exiting) return;
    setFlipped(true);
  }

  function toggleDetails() {
    if (!activeWord || !flipped || completed || exiting) return;
    const shouldExpand = !detailsExpanded;
    setDetailsExpanded(shouldExpand);
    if (shouldExpand) onEnsureWordDetails(activeWord);
  }

  function submitFeedback(status: MarkStatus) {
    if (exiting || !activeWord) return;
    onReviewWordComplete(activeWord, status);
    const nextReviewedIds = new Set(reviewedWordIdsRef.current);
    nextReviewedIds.add(activeWord.id);
    reviewedWordIdsRef.current = nextReviewedIds;
    setReviewedWordIds(nextReviewedIds);
    setReviewStatuses((current) => [...current, status]);
    setExiting(true);
    window.setTimeout(() => {
      setReviewIndex((current) => current + 1);
      setFlipped(false);
      setDetailsExpanded(false);
      setExiting(false);
    }, 240);
  }

  function continueReview() {
    const reviewedIds = reviewedWordIdsRef.current;
    const combinedExcludedWordIds = new Set([...excludedWordIds, ...reviewedIds]);
    const nextQueue = createReviewQueue(
      decks,
      deckFilter,
      reviewLimit + combinedExcludedWordIds.size,
      progress,
      mastery,
      reviewCycles,
      reviewSchedule,
      combinedExcludedWordIds,
    ).filter(
      (word) => !combinedExcludedWordIds.has(word.id),
    );
    setQueue(nextQueue.slice(0, reviewLimit));
    setReviewIndex(0);
    setFlipped(false);
    setDetailsExpanded(false);
    setExiting(false);
    setReviewStatuses([]);
    reviewStartedAtRef.current = Date.now();
  }

  useEffect(() => {
    function handleSpace(event: KeyboardEvent) {
      const target = event.target as HTMLElement | null;
      if (target?.closest("input, textarea, select, button")) return;
      if (event.code !== "Space" || completed || exiting) return;
      event.preventDefault();
      if (!flipped) flipCard();
      else toggleDetails();
    }

    window.addEventListener("keydown", handleSpace);
    return () => window.removeEventListener("keydown", handleSpace);
  }, [activeWord, completed, exiting, flipped]);

  return (
    <main className="review-page">
      <ReviewHeader
        completed={completed}
        decks={decks}
        deckFilter={deckFilter}
        onBack={onBack}
        onDeckFilterChange={changeDeckFilter}
        reviewedCount={reviewedCount}
        totalCount={queue.length}
      />
      <section className={completed ? "review-main review-main-complete" : "review-main"} aria-label="复习卡片">
        {!speechSupported && !completed ? (
          <div className="speech-support-warning">当前浏览器不支持语音播放，请用 Chrome 打开后重试。</div>
        ) : null}
        {completed ? (
          <ReviewCompletePanel
            count={queue.length}
            focusMinutes={focusMinutes}
            onBack={onBack}
            onContinueReview={continueReview}
            quickRecallPercent={quickRecallPercent}
          />
        ) : (
          activeWord && (
            <>
              <ReviewFlashcard
                detailState={getWordDetailState(activeWord)}
                detailsExpanded={detailsExpanded}
                exiting={exiting}
                flipped={flipped}
                key={activeWord.id}
                onDetailsToggle={toggleDetails}
                onFlip={flipCard}
                speechSupported={speechSupported}
                word={activeWord}
              />
              {flipped ? <ReviewFeedbackActions disabled={exiting} onFeedback={submitFeedback} /> : null}
            </>
          )
        )}
      </section>
      <ReviewFooterTip />
    </main>
  );
}

function ReviewHeader({
  completed,
  decks,
  deckFilter,
  onBack,
  onDeckFilterChange,
  reviewedCount,
  totalCount,
}: {
  completed: boolean;
  decks: VocabularyDeck[];
  deckFilter: ReviewDeckFilter;
  onBack: () => void;
  onDeckFilterChange: (deckFilter: ReviewDeckFilter) => void;
  reviewedCount: number;
  totalCount: number;
}) {
  const progressPercent = totalCount ? Math.round((reviewedCount / totalCount) * 100) : 0;
  return (
    <header className="review-header">
      <div className="review-header-inner">
        <div className="review-header-left">
          <div className="review-brand">
            <button className="review-back" type="button" onClick={onBack} aria-label="返回首页">
              <ArrowLeft size={16} />
            </button>
            <div className="wordmark" aria-label="Wordloop">
              Wordloop
            </div>
          </div>
          <span className="review-divider" />
          <div className="review-deck-select">
            <span>词库</span>
            <ReviewDeckSelect decks={decks} value={deckFilter} onValueChange={onDeckFilterChange} />
          </div>
        </div>
        <div className="review-progress-block" aria-label={`复习进度 ${reviewedCount} / ${totalCount} words`}>
          <span>
            {completed ? totalCount : reviewedCount} / {totalCount} words
          </span>
          <div className="review-progress-track" aria-hidden="true">
            <i style={{ width: `${completed ? 100 : progressPercent}%` }} />
          </div>
        </div>
      </div>
    </header>
  );
}

function ReviewFlashcard({
  detailState,
  detailsExpanded,
  exiting,
  flipped,
  onDetailsToggle,
  onFlip,
  speechSupported,
  word,
}: {
  detailState?: WordDetailsState;
  detailsExpanded: boolean;
  exiting: boolean;
  flipped: boolean;
  onDetailsToggle: () => void;
  onFlip: () => void;
  speechSupported: boolean;
  word: VocabularyWord;
}) {
  const reviewMeaning = word.meaning.replace(/^adj\.\s*/, "");
  const wordSpeechHandlers = useDelayedSpeech(word.word, 650, { lang: "en-US" });
  const exampleSpeechHandlers = useDelayedSpeech(word.example, 650, { lang: "en-US" });
  const stageClassName = [
    "review-card-stage",
    detailsExpanded ? "review-card-stage-details" : "",
    exiting ? "review-card-stage-exiting" : "",
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <div className={stageClassName}>
      <div
        className={flipped ? "review-flashcard review-flashcard-flipped" : "review-flashcard"}
        role="button"
        tabIndex={0}
        onClick={onFlip}
        onKeyDown={(event) => {
          if (event.key !== "Enter") return;
          event.preventDefault();
          onFlip();
        }}
        aria-label={flipped ? `${word.word} 释义` : "点击查看注释"}
      >
        <span className="review-flashcard-inner">
          <span className="review-card-layer" aria-hidden="true" />
          <span className="review-card-face review-card-front">
            <strong {...wordSpeechHandlers}>{word.word}</strong>
            <span>点击查看注释</span>
          </span>
          <span className="review-card-face review-card-back">
            <span className="review-word-head">
              <span>
                <strong {...wordSpeechHandlers}>{word.word}</strong>
                <span className="review-phonetic-row">
                  <i>{word.phonetic}</i>
                  <button
                    className="review-sound"
                    type="button"
                    disabled={!speechSupported}
                    aria-label={`播放 ${word.word} 发音`}
                    onClick={(event) => {
                      event.stopPropagation();
                      speakText(word.word, { lang: "en-US", notify: true });
                    }}
                  >
                    <Volume2 size={12} />
                  </button>
                </span>
              </span>
            </span>
            <span className="review-card-back-body">
              <span className="review-translation-section">
                <b>{reviewMeaning}</b>
                <span>{word.extraMeanings[0] ?? reviewMeaning}</span>
              </span>
              <span className="review-example-section">
                <span {...exampleSpeechHandlers}>
                  {word.example.split(new RegExp(`(${word.word})`, "i")).map((part, index) =>
                    part.toLowerCase() === word.word.toLowerCase() ? (
                      <em key={`${part}-${index}`}>{part}</em>
                    ) : (
                      part
                    ),
                  )}
                </span>
                <span>{word.exampleTranslation}</span>
              </span>
              <button
                className={detailsExpanded ? "review-detail-link review-detail-link-open" : "review-detail-link"}
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  onDetailsToggle();
                }}
              >
                {detailsExpanded ? "收起详情" : "查看更多"}
                {detailsExpanded ? <ChevronUp size={12} /> : <ChevronDown size={12} />}
              </button>
              {detailsExpanded ? (
                <span className="review-more-details" onClick={(event) => event.stopPropagation()}>
                  {detailState?.loading ? <span className="word-detail-status">正在生成详情...</span> : null}
                  {detailState?.error ? <span className="word-detail-status word-detail-error">{detailState.error}</span> : null}
                  {detailState?.data ? <WordAiDetailsPanel details={detailState.data} word={word} /> : null}
                </span>
              ) : null}
            </span>
          </span>
        </span>
      </div>
    </div>
  );
}

function ReviewFeedbackActions({
  disabled,
  onFeedback,
}: {
  disabled: boolean;
  onFeedback: (status: MarkStatus) => void;
}) {
  const actions: { label: string; icon: ReactNode; status: MarkStatus }[] = [
    { label: "瞬间想起", icon: <Zap size={20} />, status: "known" },
    { label: "需要想想", icon: <CircleHelp size={20} />, status: "familiar" },
    { label: "完全忘记", icon: <Ban size={20} />, status: "unknown" },
  ];

  return (
    <div className="review-feedback-actions">
      {actions.map((action) => (
        <button disabled={disabled} key={action.label} type="button" onClick={() => onFeedback(action.status)}>
          <span>{action.icon}</span>
          {action.label}
        </button>
      ))}
    </div>
  );
}

function ReviewDeckSelect({
  decks,
  onValueChange,
  value,
}: {
  decks: VocabularyDeck[];
  onValueChange: (deckFilter: ReviewDeckFilter) => void;
  value: ReviewDeckFilter;
}) {
  return (
    <Select.Root value={value} onValueChange={(nextValue) => onValueChange(nextValue as ReviewDeckFilter)}>
      <Select.Trigger className="radix-select-trigger review-select-trigger" aria-label="选择复习词库">
        <Select.Value />
        <Select.Icon asChild>
          <ChevronDown size={16} />
        </Select.Icon>
      </Select.Trigger>
      <Select.Portal>
        <Select.Content className="radix-select-content" position="popper" sideOffset={6}>
          <Select.Viewport className="radix-select-viewport">
            <Select.Item className="radix-select-item" value="all">
              <Select.ItemText>全部（默认）</Select.ItemText>
            </Select.Item>
            {decks.map((deck) => (
              <Select.Item className="radix-select-item" key={deck.id} value={deck.id}>
                <Select.ItemText>{deck.name}</Select.ItemText>
              </Select.Item>
            ))}
          </Select.Viewport>
        </Select.Content>
      </Select.Portal>
    </Select.Root>
  );
}

function ReviewCompletePanel({
  count,
  focusMinutes,
  onBack,
  onContinueReview,
  quickRecallPercent,
}: {
  count: number;
  focusMinutes: number;
  onBack: () => void;
  onContinueReview: () => void;
  quickRecallPercent: number;
}) {
  return (
    <>
      <section className="review-complete-panel">
        <span className="review-complete-icon">
          <PartyPopper size={50} />
        </span>
        <h1>太棒了！复习已完成</h1>
        <p>您今天已经复习了 {count} 个新单词，复习进度 100%。</p>
        <dl>
          <div>
            <dt>复习单词</dt>
            <dd>{count}</dd>
          </div>
          <div>
            <dt>专注时长</dt>
            <dd>{focusMinutes}m</dd>
          </div>
          <div>
            <dt>掌握率</dt>
            <dd>{quickRecallPercent}%</dd>
          </div>
        </dl>
      </section>
      <div className="review-complete-actions">
        <button className="review-outline-action" type="button" onClick={onContinueReview}>
          继续复习
        </button>
        <button className="review-primary-action" type="button" onClick={onBack}>
          完成
        </button>
      </div>
    </>
  );
}

function ReviewFooterTip() {
  return (
    <div className="review-footer-tip">
      <CircleHelp size={13} />
      <span>提示: 空格键可展开详情</span>
    </div>
  );
}

function SettingsModal({
  initialSettings,
  onClose,
  onSave,
}: {
  initialSettings: AppSettings;
  onClose: () => void;
  onSave: (settings: AppSettings) => void;
}) {
  const [draft, setDraft] = useState<AppSettings>(() => ({
    ...initialSettings,
    dailyStudyCount: Math.min(30, Math.max(5, initialSettings.dailyStudyCount)),
    dailyReviewGoal: Math.min(50, Math.max(10, initialSettings.dailyReviewGoal)),
  }));

  function updateNumber(key: "dailyStudyCount" | "dailyReviewGoal", value: number) {
    setDraft((current) => ({
      ...current,
      [key]: value,
    }));
  }

  return (
    <div className="settings-overlay" role="dialog" aria-modal="true" aria-labelledby="settings-title">
      <div className="settings-modal">
        <header className="settings-modal-header">
          <div className="settings-modal-title">
            <span>
              <Settings size={22} />
            </span>
            <h1 id="settings-title">设置</h1>
          </div>
          <button className="settings-close" type="button" onClick={onClose} aria-label="关闭设置">
            <X size={22} />
          </button>
        </header>

        <div className="settings-modal-content">
          <section className="settings-section">
            <div className="settings-section-title">
              <KeyRound size={18} />
              <h2>AI 配置</h2>
            </div>
            <label className="settings-field">
              <span>API KEY (DEEPSEEK)</span>
              <div className="settings-api-input">
                <input
                  type="text"
                  value={draft.deepseekApiKey}
                  placeholder="请输入 DeepSeek API Key"
                  onChange={(event) =>
                    setDraft((current) => ({ ...current, deepseekApiKey: event.target.value }))
                  }
                />
              </div>
            </label>
          </section>

          <section className="settings-section">
            <div className="settings-section-title">
              <Flag size={18} />
              <h2>学习目标</h2>
            </div>
            <div className="settings-goals">
              <SettingsSlider
                label="单次学习数量"
                max={30}
                min={5}
                minLabel="轻松 (5)"
                maxLabel="挑战 (30)"
                value={draft.dailyStudyCount}
                onChange={(value) => updateNumber("dailyStudyCount", value)}
              />
              <SettingsSlider
                label="每日复习目标"
                max={50}
                min={10}
                minLabel="适中 (10)"
                maxLabel="高强度 (50)"
                value={draft.dailyReviewGoal}
                onChange={(value) => updateNumber("dailyReviewGoal", value)}
              />
            </div>
          </section>
        </div>

        <footer className="settings-modal-footer">
          <button className="settings-cancel" type="button" onClick={onClose}>
            取消
          </button>
          <button className="settings-save" type="button" onClick={() => onSave(draft)}>
            保存更改
          </button>
        </footer>
      </div>
    </div>
  );
}

function SettingsSlider({
  label,
  max,
  maxLabel,
  min,
  minLabel,
  onChange,
  value,
}: {
  label: string;
  max: number;
  maxLabel: string;
  min: number;
  minLabel: string;
  onChange: (value: number) => void;
  value: number;
}) {
  return (
    <div className="settings-slider-field">
      <div className="settings-slider-head">
        <span>{label}</span>
        <strong>{value}</strong>
      </div>
      <Slider.Root
        className="settings-slider-root"
        min={min}
        max={max}
        step={1}
        value={[value]}
        onValueChange={([nextValue]) => onChange(nextValue)}
      >
        <Slider.Track className="settings-slider-track">
          <Slider.Range className="settings-slider-range" />
        </Slider.Track>
        <Slider.Thumb className="settings-slider-thumb" aria-label={label} />
      </Slider.Root>
      <div className="settings-slider-labels">
        <span>{minLabel}</span>
        <span>{maxLabel}</span>
      </div>
    </div>
  );
}

function Header({
  onOpenSettings,
  onSignOut,
  userEmail,
}: {
  onOpenSettings: () => void;
  onSignOut: () => void;
  userEmail: string;
}) {
  const avatarText = userEmail ? userEmail.slice(0, 2).toUpperCase() : "ME";
  return (
    <header className="home-header">
      <div className="brand-group">
        <div className="wordmark" aria-label="Wordloop">
          Wordloop
        </div>
        <span className="ai-badge">AI 驱动</span>
        <p>自主生成专属词库，完整闭环式单词学习工具</p>
      </div>
      <div className="header-actions">
        <button className="icon-button" type="button" aria-label="设置" onClick={onOpenSettings}>
          <Settings size={24} strokeWidth={2} />
        </button>
        <Popover.Root>
          <Popover.Trigger className="avatar" aria-label="当前用户菜单">
            {avatarText}
          </Popover.Trigger>
          <Popover.Portal>
            <Popover.Content className="user-menu-popover" align="end" sideOffset={10}>
              <div className="user-menu-info">
                <span>登录账号</span>
                <strong>{userEmail || "匿名用户"}</strong>
              </div>
              <button className="user-menu-sign-out" type="button" onClick={onSignOut}>
                退出登录
              </button>
            </Popover.Content>
          </Popover.Portal>
        </Popover.Root>
      </div>
    </header>
  );
}

function StudyHeader({
  canEnterQuiz,
  deck,
  markedCount,
  onAddTenWords,
  onBack,
  onEnterQuiz,
  totalCount,
}: {
  canEnterQuiz: boolean;
  deck: VocabularyDeck;
  markedCount: number;
  onAddTenWords: () => void;
  onBack: () => void;
  onEnterQuiz: () => void;
  totalCount: number;
}) {
  return (
    <header className="study-header">
      <div className="study-header-inner">
        <div className="study-brand">
          <button className="back-button" type="button" onClick={onBack} aria-label="返回首页">
            <ArrowLeft size={18} />
          </button>
          <div className="wordmark" aria-label="Wordloop">
            Wordloop
          </div>
          <span className="study-divider" />
          <span className="study-deck-name">{deck.name}</span>
        </div>
        <div className="study-header-actions">
          <span className="study-progress-pill">
            {markedCount}/{totalCount}
          </span>
          <AppButton tone="neutral" icon={<Plus size={24} />} onClick={onAddTenWords}>
            再学10个
          </AppButton>
          <AppButton
            tone="primary"
            disabled={!canEnterQuiz}
            icon={<MoveRight size={24} />}
            onClick={onEnterQuiz}
          >
            立即测验
          </AppButton>
        </div>
      </div>
    </header>
  );
}

function DeckCard({
  deck,
  isMenuOpen,
  learningComplete,
  masteredWords,
  onConfirmAction,
  onExport,
  onExpand,
  onImportWords,
  onRename,
  onToggleMenu,
  onStartStudy,
  progress,
  reviewCycles,
  reviewSchedule,
}: {
  deck: VocabularyDeck;
  isMenuOpen: boolean;
  learningComplete: boolean;
  masteredWords: Record<string, true>;
  onConfirmAction: (deck: VocabularyDeck, action: DeckConfirmAction) => void;
  onExport: (deck: VocabularyDeck) => void;
  onExpand: (deckId: string) => void;
  onImportWords: (deckId: string) => void;
  onRename: (deckId: string) => void;
  onToggleMenu: (deckId: string) => void;
  onStartStudy: (deckId: string) => void;
  progress: Record<string, MarkStatus>;
  reviewCycles: Record<string, number>;
  reviewSchedule: Record<string, StoredReviewSchedule[string][string]>;
}) {
  const learned = Object.keys(progress).length;
  const mastered = deck.words.filter((word) => {
    const memory = reviewSchedule[word.id] ?? getFallbackReviewMemory(progress[word.id], reviewCycles[word.id] ?? 0);
    return masteredWords[word.id] || isMemoryMastered(memory);
  }).length;
  const reviewing = deck.words.filter((word) => {
    if (!progress[word.id]) return false;
    const memory = reviewSchedule[word.id] ?? getFallbackReviewMemory(progress[word.id], reviewCycles[word.id] ?? 0);
    return !isMemoryMastered(memory) && isReviewDue(memory);
  }).length;
  const progressPercent = deck.words.length ? Math.round((learned / deck.words.length) * 100) : 0;
  const sourceLabel = getDeckSourceLabel(deck);
  const stats = [
    { label: "已学/总数", value: `${learned}/${deck.words.length}`, tone: "normal" },
    { label: "复习中", value: reviewing, tone: "accent" },
    { label: "已掌握", value: mastered, tone: "normal" },
  ];

  return (
    <article className={learningComplete ? "deck-card deck-card-complete" : "deck-card"}>
      <div className="deck-title-row">
        <div className="deck-title-main">
          <h3>{deck.name}</h3>
          <span className={`deck-source-tag deck-source-tag-${sourceLabel === "AI 生成" ? "ai" : sourceLabel === "内置" ? "builtin" : "local"}`}>
            {sourceLabel}
          </span>
        </div>
        <button
          className="more-button"
          type="button"
          aria-expanded={isMenuOpen}
          aria-label={`${deck.name} 更多`}
          onClick={() => onToggleMenu(deck.id)}
        >
          <CircleEllipsis size={20} />
        </button>
        {isMenuOpen ? (
          <DeckMoreMenu
            deck={deck}
            onConfirmAction={onConfirmAction}
            onExport={onExport}
            onExpand={onExpand}
            onImportWords={onImportWords}
            onRename={onRename}
          />
        ) : null}
      </div>

      <div className="progress-block">
        <div className="progress-label-row">
          <span>进度</span>
          <span>{progressPercent}%</span>
        </div>
        <div className="progress-track" aria-hidden="true">
          <span style={{ width: `${progressPercent}%` }} />
        </div>
      </div>

      <dl className="deck-stats">
        {stats.map((stat) => (
          <div key={stat.label}>
            <dt>{stat.label}</dt>
            <dd className={stat.tone === "accent" ? "accent-value" : undefined}>{stat.value}</dd>
          </div>
        ))}
      </dl>

      <AppButton tone="soft" fullWidth disabled={learningComplete} onClick={() => onStartStudy(deck.id)}>
        {learningComplete ? "已学完" : "立即学习"}
      </AppButton>
    </article>
  );
}

function DeckMoreMenu({
  deck,
  onConfirmAction,
  onExport,
  onExpand,
  onImportWords,
  onRename,
}: {
  deck: VocabularyDeck;
  onConfirmAction: (deck: VocabularyDeck, action: DeckConfirmAction) => void;
  onExport: (deck: VocabularyDeck) => void;
  onExpand: (deckId: string) => void;
  onImportWords: (deckId: string) => void;
  onRename: (deckId: string) => void;
}) {
  const supportsAiExpansion = Boolean(deck.aiGenerated || deck.description.includes("AI 生词"));
  const editActions = [
    { label: "重命名", icon: <Pencil size={16} />, tone: "normal", onClick: () => onRename(deck.id) },
    { label: "导入单词", icon: <FileInput size={16} />, tone: "normal", onClick: () => onImportWords(deck.id) },
    ...(supportsAiExpansion ? [{ label: "AI 扩充词库", icon: <Bot size={16} />, tone: "normal", onClick: () => onExpand(deck.id) }] : []),
  ];
  const dataActions = [
    { label: "导出excel", icon: <Download size={16} />, tone: "normal", onClick: () => onExport(deck) },
    { label: "重置学习记录", icon: <RefreshCcw size={16} />, tone: "danger", onClick: () => onConfirmAction(deck, "reset") },
    { label: "删除词库", icon: <Trash2 size={16} />, tone: "danger", onClick: () => onConfirmAction(deck, "delete") },
  ];

  return (
    <div className="deck-more-menu" role="menu">
      {editActions.map((action) => (
        <button
          className={`deck-more-item deck-more-item-${action.tone}`}
          key={action.label}
          type="button"
          role="menuitem"
          onClick={action.onClick}
        >
          {action.icon}
          {action.label}
        </button>
      ))}
      {dataActions.map((action) => (
        <button
          className={`deck-more-item deck-more-item-${action.tone}`}
          key={action.label}
          type="button"
          role="menuitem"
          onClick={action.onClick}
        >
          {action.icon}
          {action.label}
        </button>
      ))}
    </div>
  );
}

function RenameDeckDialog({
  deck,
  existingNames,
  onClose,
  onSave,
}: {
  deck: VocabularyDeck;
  existingNames: string[];
  onClose: () => void;
  onSave: (deckId: string, nextName: string) => void;
}) {
  const [name, setName] = useState(deck.name);
  const trimmedName = name.trim();
  const hasDuplicateName = existingNames.includes(trimmedName);
  const canSave = Boolean(trimmedName) && trimmedName !== deck.name && !hasDuplicateName;

  function submitRename(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!canSave) return;
    onSave(deck.id, trimmedName);
  }

  return (
    <Dialog.Root open onOpenChange={(open) => (!open ? onClose() : undefined)}>
      <Dialog.Portal>
        <Dialog.Overlay className="rename-dialog-overlay" />
        <Dialog.Content className="rename-dialog-content">
          <div className="rename-dialog-header">
            <div>
              <Dialog.Title className="rename-dialog-title">重命名词库</Dialog.Title>
              <Dialog.Description className="rename-dialog-description">修改后会同步更新首页词库卡片名称。</Dialog.Description>
            </div>
            <Dialog.Close className="rename-dialog-close" aria-label="关闭">
              <X size={18} />
            </Dialog.Close>
          </div>

          <form className="rename-dialog-form" onSubmit={submitRename}>
            <label className="rename-dialog-field" aria-label="词库名称">
              <input
                autoFocus
                maxLength={24}
                value={name}
                onChange={(event) => setName(event.target.value)}
              />
            </label>
            {hasDuplicateName ? <p className="rename-dialog-error">已存在同名词库，请换一个名称。</p> : null}
            <div className="rename-dialog-footer">
              <Dialog.Close className="modal-cancel-button" type="button">
                取消
              </Dialog.Close>
              <button className="modal-primary-button" type="submit" disabled={!canSave}>
                保存
              </button>
            </div>
          </form>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function ExpandDeckDialog({
  apiKey,
  deck,
  onClose,
  onExpand,
}: {
  apiKey: string;
  deck: VocabularyDeck | null;
  onClose: () => void;
  onExpand: (deck: VocabularyDeck, words: ImportedWordInput[]) => void;
}) {
  const currentCount = deck?.words.length ?? 0;
  const maxCount = 300;
  const initialTargetCount = Math.min(maxCount, Math.max(currentCount + 20, currentCount));
  const [targetCount, setTargetCount] = useState(initialTargetCount);
  const [notice, setNotice] = useState("");
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    if (!deck) return;
    setTargetCount(Math.min(maxCount, Math.max(deck.words.length + 20, deck.words.length)));
    setNotice("");
  }, [deck?.id]);

  if (!deck) return null;

  const extraCount = Math.max(0, targetCount - currentCount);
  const canExpand = Boolean(apiKey.trim()) && extraCount > 0 && !submitting;

  async function confirmExpand() {
    if (!deck || !extraCount || submitting) return;
    if (!apiKey.trim()) {
      setNotice("请先在设置里配置 DeepSeek API Key。");
      return;
    }
    setSubmitting(true);
    setNotice("正在用 AI 扩充词库...");
    try {
      const words = await generateDeckWordsWithDeepSeek(apiKey.trim(), deck.name, extraCount);
      if (!words.length) {
        setNotice("AI 未生成可导入的单词，请调整数量后重试。");
        return;
      }
      onExpand(deck, words);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : "AI 扩充失败，请稍后重试。");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <Dialog.Root open onOpenChange={(open) => (!open ? onClose() : undefined)}>
      <Dialog.Portal>
        <Dialog.Overlay className="confirm-dialog-overlay" />
        <Dialog.Content className="expand-dialog-content">
          <div className="expand-dialog-header">
            <div>
              <Dialog.Title className="expand-dialog-title">AI 扩充词库</Dialog.Title>
              <Dialog.Description className="expand-dialog-description">
                扩充「{deck.name}」到目标词汇量，AI 会补齐释义和例句。
              </Dialog.Description>
            </div>
            <Dialog.Close className="rename-dialog-close" aria-label="关闭">
              <X size={18} />
            </Dialog.Close>
          </div>

          <div className="expand-dialog-body">
            <div className="expand-count-row">
              <span>目标词汇量</span>
              <strong>{targetCount} 词</strong>
            </div>
            <Slider.Root
              className="import-slider-root"
              min={currentCount}
              max={maxCount}
              step={5}
              value={[targetCount]}
              onValueChange={([nextValue]) => setTargetCount(nextValue)}
            >
              <Slider.Track className="import-slider-track">
                <Slider.Range className="import-slider-range" />
              </Slider.Track>
              <Slider.Thumb className="import-slider-thumb" aria-label="目标词汇量" />
            </Slider.Root>
            <div className="expand-slider-labels">
              <span>{currentCount}</span>
              <span>{maxCount}</span>
            </div>
            <p className="expand-dialog-helper">
              {currentCount >= maxCount ? "当前词库已达到扩充上限。" : `本次预计新增 ${extraCount} 个单词。`}
            </p>
            {notice ? <p className="import-notice">{notice}</p> : null}
          </div>

          <div className="expand-dialog-footer">
            <Dialog.Close className="modal-cancel-button" type="button">
              取消
            </Dialog.Close>
            <button className="modal-primary-button" type="button" disabled={!canExpand} onClick={() => void confirmExpand()}>
              {submitting ? "扩充中" : "开始扩充"}
            </button>
          </div>
        </Dialog.Content>
      </Dialog.Portal>
    </Dialog.Root>
  );
}

function DeckConfirmDialog({
  action,
  deck,
  onClose,
  onConfirm,
}: {
  action: DeckConfirmAction;
  deck: VocabularyDeck;
  onClose: () => void;
  onConfirm: () => void;
}) {
  const isDelete = action === "delete";
  const title = isDelete ? "删除词库" : "重置词库";
  const description = isDelete
    ? deck.custom
      ? `删除「${deck.name}」后，该词库和学习记录将从本地移除。`
      : `删除「${deck.name}」后，该内置词库将从首页隐藏，学习记录也会清空。`
    : `重置「${deck.name}」后，该词库的学习进度和掌握记录将被清空。`;
  const actionLabel = isDelete ? "确认删除" : "确认重置";

  return (
    <AlertDialog.Root open onOpenChange={(open) => (!open ? onClose() : undefined)}>
      <AlertDialog.Portal>
        <AlertDialog.Overlay className="confirm-dialog-overlay" />
        <AlertDialog.Content className="confirm-dialog-content">
          <AlertDialog.Title className="confirm-dialog-title">{title}</AlertDialog.Title>
          <AlertDialog.Description className="confirm-dialog-description">{description}</AlertDialog.Description>
          <div className="confirm-dialog-footer">
            <AlertDialog.Cancel className="modal-cancel-button">取消</AlertDialog.Cancel>
            <AlertDialog.Action
              className="modal-danger-button"
              onClick={onConfirm}
            >
              {actionLabel}
            </AlertDialog.Action>
          </div>
        </AlertDialog.Content>
      </AlertDialog.Portal>
    </AlertDialog.Root>
  );
}

function AppToast({ message }: { message: string }) {
  return (
    <div className="app-toast" role="status" aria-live="polite">
      <CircleCheck size={18} />
      <span>{message}</span>
    </div>
  );
}

function AuthStatusPage({ message }: { message: string }) {
  return (
    <main className="auth-page">
      <section className="auth-card">
        <div className="wordmark" aria-label="Wordloop">
          Wordloop
        </div>
        <p>{message}</p>
      </section>
    </main>
  );
}

function EmailLoginPage({
  onAnonymousLogin,
  onSubmit,
}: {
  onAnonymousLogin: () => Promise<void>;
  onSubmit: (email: string, password: string, mode: AuthMode) => Promise<void>;
}) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [mode, setMode] = useState<AuthMode>("signIn");
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");

  async function submitLogin(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!email.trim() || password.length < 6 || submitting) return;
    setSubmitting(true);
    setError("");
    try {
      await onSubmit(email.trim(), password, mode);
    } catch (error) {
      console.error("Password auth failed", error);
      const message = error instanceof Error ? error.message : "未知错误";
      const friendlyMessage =
        message === "Invalid login credentials"
          ? "邮箱或密码不正确；如果刚注册，请先到邮箱里确认账号，或在 Supabase 里关闭 Confirm email。"
          : message;
      setError(`${mode === "signIn" ? "登录" : "注册"}失败：${friendlyMessage}`);
    } finally {
      setSubmitting(false);
    }
  }

  async function submitAnonymousLogin() {
    if (submitting) return;
    setSubmitting(true);
    setError("");
    try {
      await onAnonymousLogin();
    } catch (error) {
      console.error("Anonymous login failed", error);
      const message = error instanceof Error ? error.message : "未知错误";
      setError(`匿名登录失败：${message}`);
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <main className="auth-page">
      <form className="auth-card" onSubmit={submitLogin}>
        <div className="wordmark" aria-label="Wordloop">
          Wordloop
        </div>
        <h1>{mode === "signIn" ? "邮箱密码登录" : "创建账号"}</h1>
        <p>{mode === "signIn" ? "输入邮箱和密码即可进入应用。" : "首次使用请创建一个邮箱密码账号。"}</p>
        <label>
          <span>邮箱</span>
          <input
            autoFocus
            type="email"
            value={email}
            placeholder="you@example.com"
            onChange={(event) => setEmail(event.target.value)}
          />
        </label>
        <label>
          <span>密码</span>
          <input
            type="password"
            value={password}
            placeholder="至少 6 位"
            onChange={(event) => setPassword(event.target.value)}
          />
        </label>
        {error ? <strong>{error}</strong> : null}
        <button type="submit" disabled={!email.trim() || password.length < 6 || submitting}>
          {submitting ? "处理中" : mode === "signIn" ? "登录" : "注册并登录"}
        </button>
        <button className="auth-secondary-button" type="button" disabled={submitting} onClick={submitAnonymousLogin}>
          匿名登录，先进入应用
        </button>
        <button
          className="auth-switch-button"
          type="button"
          onClick={() => {
            setMode((current) => (current === "signIn" ? "signUp" : "signIn"));
            setError("");
          }}
        >
          {mode === "signIn" ? "没有账号？创建一个" : "已有账号？去登录"}
        </button>
      </form>
    </main>
  );
}

function QuickWordModal({
  apiKey,
  decks,
  onClose,
  onCreateDeck,
  onSave,
}: {
  apiKey: string;
  decks: VocabularyDeck[];
  onClose: () => void;
  onCreateDeck: (name: string) => string;
  onSave: (words: Array<string | ImportedWordInput>, deckName: string) => void;
}) {
  const [input, setInput] = useState("");
  const [selectedDeckName, setSelectedDeckName] = useState("速记单词");
  const [newDeckName, setNewDeckName] = useState("");
  const [createDeckOpen, setCreateDeckOpen] = useState(false);
  const [notice, setNotice] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const words = parseWordInput(input);
  const deckNames = Array.from(new Set(["速记单词", ...decks.map((deck) => deck.name)]));

  function confirmCreateDeck() {
    const createdName = onCreateDeck(newDeckName);
    if (!createdName) return;
    setSelectedDeckName(createdName);
    setNewDeckName("");
    setCreateDeckOpen(false);
  }

  async function confirmSave() {
    if (!words.length || submitting) return;
    setSubmitting(true);
    setNotice(apiKey.trim() ? "正在用 AI 补齐释义和例句..." : "未配置 API Key，将直接导入纯单词。");
    try {
      const completedWords = await completeImportedWordsWithDeepSeek(apiKey, words);
      onSave(completedWords, selectedDeckName);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : "AI 补齐失败，请稍后重试。");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="home-modal-overlay" role="dialog" aria-modal="true" aria-labelledby="quick-modal-title">
      <section className="quick-modal">
        <header className="home-modal-header">
          <h2 id="quick-modal-title">
            <Zap size={24} />
            快速记单词
          </h2>
          <button type="button" onClick={onClose} aria-label="关闭">
            <X size={14} />
          </button>
        </header>
        <div className="quick-modal-body">
          <label className="home-modal-field">
            <span>输入单词</span>
            <textarea
              value={input}
              placeholder="多个单词可使用空格、逗号、分号等隔开"
              onChange={(event) => setInput(event.target.value)}
            />
          </label>
          <label className="home-modal-field">
            <span className="home-modal-field-head">
              <span>选择目标词库</span>
              <Popover.Root open={createDeckOpen} onOpenChange={setCreateDeckOpen}>
                <Popover.Trigger className="quick-create-link" type="button">
                  新建词库
                </Popover.Trigger>
                <Popover.Portal>
                  <Popover.Content className="quick-create-popover" align="end" sideOffset={8}>
                    <label>
                      <span>词库名称</span>
                      <input
                        autoFocus
                        value={newDeckName}
                        placeholder="输入新词库名称"
                        onChange={(event) => setNewDeckName(event.target.value)}
                        onKeyDown={(event) => {
                          if (event.key === "Enter") {
                            event.preventDefault();
                            confirmCreateDeck();
                          }
                        }}
                      />
                    </label>
                    <div>
                      <button type="button" onClick={() => setCreateDeckOpen(false)}>
                        取消
                      </button>
                      <button type="button" disabled={!newDeckName.trim()} onClick={confirmCreateDeck}>
                        确认
                      </button>
                    </div>
                    <Popover.Arrow className="quick-create-arrow" />
                  </Popover.Content>
                </Popover.Portal>
              </Popover.Root>
            </span>
            <DeckNameSelect
              names={deckNames}
              value={selectedDeckName}
              onValueChange={setSelectedDeckName}
            />
          </label>
          {notice ? <p className="import-notice">{notice}</p> : null}
        </div>
        <footer className="home-modal-footer">
          <button className="modal-cancel-button" type="button" onClick={onClose}>
            取消
          </button>
          <button
            className="modal-primary-button"
            type="button"
            disabled={!words.length || submitting}
            onClick={() => void confirmSave()}
          >
            {submitting ? "补齐中" : "确认并导入"}
          </button>
        </footer>
      </section>
    </div>
  );
}

function DeckNameSelect({
  names,
  onValueChange,
  value,
}: {
  names: string[];
  onValueChange: (value: string) => void;
  value: string;
}) {
  return (
    <Select.Root value={value} onValueChange={onValueChange}>
      <Select.Trigger className="radix-select-trigger quick-deck-trigger" aria-label="选择目标词库">
        <Select.Value />
        <Select.Icon asChild>
          <ChevronDown size={16} />
        </Select.Icon>
      </Select.Trigger>
      <Select.Portal>
        <Select.Content className="radix-select-content" position="popper" sideOffset={6}>
          <Select.Viewport className="radix-select-viewport">
            {names.map((name) => (
              <Select.Item className="radix-select-item" key={name} value={name}>
                <Select.ItemText>{name}</Select.ItemText>
              </Select.Item>
            ))}
          </Select.Viewport>
        </Select.Content>
      </Select.Portal>
    </Select.Root>
  );
}

function ImportDeckModal({
  activeTab,
  apiKey,
  onClose,
  onCreateDeck,
  onTabChange,
}: {
  activeTab: ImportDeckTab;
  apiKey: string;
  onClose: () => void;
  onCreateDeck: (name: string, words: Array<string | ImportedWordInput>, description?: string, aiGenerated?: boolean) => void;
  onTabChange: (tab: ImportDeckTab) => void;
}) {
  const [topic, setTopic] = useState("");
  const [wordCount, setWordCount] = useState(50);
  const [fileWords, setFileWords] = useState<ImportedWordInput[]>([]);
  const [fileDeckName, setFileDeckName] = useState("");
  const [notice, setNotice] = useState("");
  const [submitting, setSubmitting] = useState(false);

  async function readImportFile(file: File | null) {
    if (!file) return;
    const words = (await parseWordsFromFile(file)).slice(0, 300);
    if (!words.length) {
      setNotice("未识别到可导入的单词。");
      return;
    }
    setFileDeckName(file.name.replace(/\.[^.]+$/, "") || "文件导入词库");
    setFileWords(words);
    setNotice(`已识别 ${words.length} 个单词，点击确认导入。`);
  }

  async function generateWithAi() {
    const cleanTopic = topic.trim();
    if (!cleanTopic) {
      setNotice("请输入词库主题。");
      return;
    }
    setSubmitting(true);
    setNotice("");
    try {
      const words = apiKey.trim()
        ? await generateDeckWordsWithDeepSeek(apiKey.trim(), cleanTopic, wordCount)
        : Array.from({ length: wordCount }, (_, index) => ({ word: `${cleanTopic} ${index + 1}` }));
      if (!words.length) {
        setNotice("未生成可导入的单词。");
        return;
      }
      onCreateDeck(`${cleanTopic}词库`, words, "通过 AI 生词入口创建的本地词库。", true);
      onClose();
    } catch (error) {
      setNotice(error instanceof Error ? error.message : "AI 生词失败，请稍后重试。");
    } finally {
      setSubmitting(false);
    }
  }

  function confirmFileImport() {
    void confirmFileImportAsync();
  }

  async function confirmFileImportAsync() {
    if (!fileWords.length) {
      setNotice("请先选择要导入的词库文件。");
      return;
    }
    setSubmitting(true);
    setNotice(
      fileWords.some(needsWordCompletion)
        ? apiKey.trim()
          ? "正在用 AI 补齐释义和例句..."
          : "未配置 API Key，将直接导入纯单词。"
        : "",
    );
    try {
      const completedWords = await completeImportedWordsWithDeepSeek(apiKey, fileWords);
      onCreateDeck(fileDeckName || "文件导入词库", completedWords, "通过文件上传创建的词库。");
      onClose();
    } catch (error) {
      setNotice(error instanceof Error ? error.message : "AI 补齐失败，请稍后重试。");
    } finally {
      setSubmitting(false);
    }
  }

  function handleConfirm() {
    if (activeTab === "ai") void generateWithAi();
    else confirmFileImport();
  }

  return (
    <div className="home-modal-overlay" role="dialog" aria-modal="true" aria-labelledby="import-modal-title">
      <section className="import-modal">
        <header className="import-modal-header">
          <h2 id="import-modal-title">
            <Import size={25} />
            导入新词库
          </h2>
          <button type="button" onClick={onClose} aria-label="关闭">
            <X size={14} />
          </button>
        </header>
        <div className={activeTab === "file" ? "import-shell import-shell-file" : "import-shell"}>
          <aside className="import-side-nav" role="tablist">
            <button
              className={activeTab === "ai" ? "import-side-tab import-side-tab-active" : "import-side-tab"}
              type="button"
              onClick={() => onTabChange("ai")}
            >
              <Sparkles size={22} />
              AI 智能生成
            </button>
            <button
              className={activeTab === "file" ? "import-side-tab import-side-tab-active" : "import-side-tab"}
              type="button"
              onClick={() => onTabChange("file")}
            >
              <FileInput size={20} />
              文件导入
            </button>
          </aside>
          <div className="import-main">
            <div className="import-modal-body">
              {activeTab === "ai" ? (
                <section className="import-ai-tab">
                  <label className="import-field">
                    <span>词库主题或主题词</span>
                    <input
                      value={topic}
                      placeholder="例如：法律、建筑、托福口语高频词..."
                      onChange={(event) => setTopic(event.target.value)}
                    />
                  </label>
                  <label className="import-count-field">
                    <span>
                      <span>生成词汇量</span>
                      <strong>{wordCount} 词</strong>
                    </span>
                    <Slider.Root
                      className="import-slider-root"
                      min={10}
                      max={300}
                      step={5}
                      value={[wordCount]}
                      onValueChange={([nextValue]) => setWordCount(nextValue)}
                    >
                      <Slider.Track className="import-slider-track">
                        <Slider.Range className="import-slider-range" />
                      </Slider.Track>
                      <Slider.Thumb className="import-slider-thumb" aria-label="生成词汇量" />
                    </Slider.Root>
                  </label>
                  <p className="import-tip">
                    <Info size={14} />
                    AI 将根据您的学习目标，智能推荐最核心的词汇量。
                  </p>
                </section>
              ) : (
                <section className="import-file-tab">
                  <label
                    className="upload-dropzone"
                    onDragOver={(event) => event.preventDefault()}
                    onDrop={(event) => {
                      event.preventDefault();
                      void readImportFile(event.dataTransfer.files?.[0] ?? null);
                    }}
                  >
                    <input
                      accept=".txt,.csv,.xls,.xlsx"
                      type="file"
                      onChange={(event) => void readImportFile(event.target.files?.[0] ?? null)}
                    />
                    <span>
                      <Import size={44} />
                    </span>
                    <strong>{fileWords.length ? `已识别 ${fileWords.length} 个单词` : "拖拽或点击上传"}</strong>
                    <p>支持 .txt, .csv, .xls, .xlsx 格式文件。若包含释义或例句将一并导入，缺失内容将由 AI 智能补全。</p>
                  </label>
                  <p className="import-tip">
                    <Info size={14} />
                    温馨提示：上传单词列表后，我们的 AI 会为您自动匹配音标与权威释义。
                  </p>
                </section>
              )}
              {notice ? <p className="import-notice">{notice}</p> : null}
            </div>
            <footer className="home-modal-footer import-modal-footer">
              <button className="modal-cancel-button" type="button" onClick={onClose}>
                取消
              </button>
              <button
                className="modal-primary-button"
                type="button"
                disabled={submitting || (activeTab === "ai" ? !topic.trim() : !fileWords.length)}
                onClick={handleConfirm}
              >
                {submitting ? (
                  "生成中"
                ) : activeTab === "ai" ? (
                  <>
                    <Sparkles size={18} />
                    生成并导入
                  </>
                ) : (
                  "确认导入"
                )}
              </button>
            </footer>
          </div>
        </div>
      </section>
    </div>
  );
}

function ImportWordsModal({
  apiKey,
  deck,
  onClose,
  onSave,
}: {
  apiKey: string;
  deck: VocabularyDeck | null;
  onClose: () => void;
  onSave: (deck: VocabularyDeck, words: Array<string | ImportedWordInput>) => void;
}) {
  const [input, setInput] = useState("");
  const [fileWords, setFileWords] = useState<ImportedWordInput[]>([]);
  const [notice, setNotice] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const words = Array.from(
    new Map(
      [...parseWordInput(input).map((word) => ({ word })), ...fileWords].map((item) => [item.word.trim().toLowerCase(), item]),
    ).values(),
  );

  async function readFile(file: File | null) {
    if (!file) return;
    const parsedWords = (await parseWordsFromFile(file)).slice(0, 300);
    if (!parsedWords.length) {
      setNotice("未识别到可导入的单词。");
      return;
    }
    setFileWords(parsedWords);
    setNotice(`已识别 ${parsedWords.length} 个单词。`);
  }

  async function confirmImport() {
    if (!deck || !words.length) return;
    setSubmitting(true);
    setNotice(
      words.some(needsWordCompletion)
        ? apiKey.trim()
          ? "正在用 AI 补齐释义和例句..."
          : "未配置 API Key，将直接导入纯单词。"
        : "",
    );
    try {
      const completedWords = await completeImportedWordsWithDeepSeek(apiKey, words);
      onSave(deck, completedWords);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : "AI 补齐失败，请稍后重试。");
    } finally {
      setSubmitting(false);
    }
  }

  if (!deck) return null;

  return (
    <div className="home-modal-overlay" role="dialog" aria-modal="true" aria-labelledby="import-words-title">
      <section className="import-words-modal">
        <header className="import-words-header">
          <div>
            <h2 id="import-words-title">导入单词</h2>
            <p>单词将导入至「{deck.name}」词库</p>
          </div>
          <button type="button" onClick={onClose} aria-label="关闭">
            <X size={14} />
          </button>
        </header>
        <div className="import-words-body">
          <label className="import-words-field">
            <span>请输入单词或上传文件</span>
            <textarea
              value={input}
              placeholder="多个单词可使用空格、逗号、分号等隔开"
              onChange={(event) => setInput(event.target.value)}
            />
          </label>
          <label className="import-words-upload">
            <input
              accept=".txt,.csv,.xls,.xlsx"
              type="file"
              onChange={(event) => void readFile(event.target.files?.[0] ?? null)}
            />
            <strong>＋ 上传文件</strong>
            <span>支持 .txt .csv .xls .xlsx 格式文件。若包含释义或例句将一并导入，缺失内容将由 AI 智能补全。</span>
          </label>
          {notice ? <p className="import-notice">{notice}</p> : null}
        </div>
        <footer className="home-modal-footer import-words-footer">
          <button className="modal-cancel-button" type="button" onClick={onClose}>
            取消
          </button>
          <button className="modal-primary-button" type="button" disabled={!words.length || submitting} onClick={() => void confirmImport()}>
            {submitting ? "补齐中" : "确认并导入"}
          </button>
        </footer>
      </section>
    </div>
  );
}

function ClassicDeckPanel() {
  const examDecks = [
    { title: "英语四六级", subtitle: "CET-4 / CET-6 核心词", icon: <GraduationCap size={18} /> },
    { title: "雅思 / 托福", subtitle: "IELTS & TOEFL 必备", icon: <Plane size={18} /> },
    { title: "专四 / 专八", subtitle: "英语专业高阶词汇", icon: <BookOpen size={18} /> },
    { title: "小学 / 初中英语", subtitle: "基础教育大纲词汇", icon: <CircleCheck size={18} /> },
    { title: "新概念英语", subtitle: "经典教材全系列", icon: <FileText size={18} /> },
  ];
  const topics = [
    { label: "出国旅游", icon: <Plane size={20} /> },
    { label: "职场商务", icon: <Briefcase size={20} /> },
    { label: "餐饮文化", icon: <Utensils size={20} /> },
    { label: "科技新闻", icon: <Newspaper size={20} /> },
    { label: "医学专业", icon: <Stethoscope size={20} /> },
  ];

  return (
    <div className="classic-panel">
      <h3>考试相关</h3>
      <div className="classic-grid">
        {examDecks.map((deck) => (
          <button key={deck.title} type="button">
            <span>{deck.icon}</span>
            <strong>{deck.title}</strong>
            <small>{deck.subtitle}</small>
            <Plus size={14} />
          </button>
        ))}
      </div>
      <h3>主题分类</h3>
      <div className="classic-topic-row">
        {topics.map((topic) => (
          <button key={topic.label} type="button">
            {topic.icon}
            <span>{topic.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

function StudyPage({
  deck,
  expandedWordIds,
  onAddTenWords,
  onBack,
  onEnterQuiz,
  onMarkWord,
  onToggleExpanded,
  onUndoWord,
  quizNotice,
  session,
  speechSupported,
  wordDetailsById,
}: {
  deck: VocabularyDeck;
  expandedWordIds: Set<string>;
  onAddTenWords: () => void;
  onBack: () => void;
  onEnterQuiz: () => void;
  onMarkWord: (wordId: string, status: MarkStatus) => void;
  onToggleExpanded: (word: VocabularyWord) => void;
  onUndoWord: (wordId: string) => void;
  quizNotice: string;
  session: StudySession;
  speechSupported: boolean;
  wordDetailsById: Record<string, WordDetailsState>;
}) {
  const sessionWords = useMemo(
    () => session.wordIds.map((id) => deck.words.find((word) => word.id === id)).filter(Boolean),
    [deck.words, session.wordIds],
  ) as VocabularyWord[];
  const pendingWords = sessionWords.filter((word) => !session.statuses[word.id]);
  const markedCount = sessionWords.length - pendingWords.length;
  const canEnterQuiz = markedCount > 0;

  return (
    <main className="study-page">
      <StudyHeader
        canEnterQuiz={canEnterQuiz}
        deck={deck}
        markedCount={markedCount}
        onAddTenWords={onAddTenWords}
        onBack={onBack}
        onEnterQuiz={onEnterQuiz}
        totalCount={sessionWords.length}
      />
      <div className="study-main">
        <section className="study-list-section" aria-labelledby="study-list-title">
          <h1 id="study-list-title" className="visually-hidden">
            本次待学习，还剩 {pendingWords.length} 个
          </h1>

          {quizNotice ? <div className="quiz-notice">{quizNotice}</div> : null}
          {!speechSupported ? (
            <div className="speech-support-warning">当前浏览器不支持语音播放，请用 Chrome 打开后重试。</div>
          ) : null}

          {pendingWords.length ? (
            <div className="word-card-list">
              {pendingWords.map((word) => (
                <WordCard
                  detailState={wordDetailsById[getWordDetailKey(deck.id, word.id)]}
                  expanded={expandedWordIds.has(word.id)}
                  key={word.id}
                  onMarkWord={onMarkWord}
                  onToggleExpanded={onToggleExpanded}
                  speechSupported={speechSupported}
                  word={word}
                />
              ))}
            </div>
          ) : (
            <StudyCompleteCard
              count={sessionWords.length}
              onAddTenWords={onAddTenWords}
              onEnterQuiz={onEnterQuiz}
            />
          )}
        </section>

        <MarkedSidebar onUndoWord={onUndoWord} session={session} words={sessionWords} />
      </div>
    </main>
  );
}

function WordCard({
  detailState,
  expanded,
  onMarkWord,
  onToggleExpanded,
  speechSupported,
  word,
}: {
  detailState?: WordDetailsState;
  expanded: boolean;
  onMarkWord: (wordId: string, status: MarkStatus) => void;
  onToggleExpanded: (word: VocabularyWord) => void;
  speechSupported: boolean;
  word: VocabularyWord;
}) {
  const [markingStatus, setMarkingStatus] = useState<MarkStatus | null>(null);
  const markTimeoutRef = useRef<number | null>(null);
  const wordSpeechHandlers = useDelayedSpeech(word.word, 650, { lang: "en-US" });
  const exampleSpeechHandlers = useDelayedSpeech(word.example, 650, { lang: "en-US" });

  useEffect(() => {
    return () => {
      if (markTimeoutRef.current) window.clearTimeout(markTimeoutRef.current);
    };
  }, []);

  function requestMark(status: MarkStatus) {
    if (markingStatus) return;
    setMarkingStatus(status);
    markTimeoutRef.current = window.setTimeout(() => {
      onMarkWord(word.id, status);
    }, 170);
  }

  const cardClassName = [
    "word-card",
    expanded ? "word-card-expanded" : "",
    markingStatus ? `word-card-exiting word-card-exiting-${markingStatus}` : "",
  ]
    .filter(Boolean)
    .join(" ");

  return (
    <article className={cardClassName}>
      <div className="word-card-body">
        <div className="word-card-top">
          <div>
            <h2 {...wordSpeechHandlers}>{word.word}</h2>
            <p className="phonetic">
              {word.phonetic}
              <button
                className="word-sound-button"
                type="button"
                disabled={!speechSupported}
                aria-label={`播放 ${word.word} 发音`}
                onClick={() => speakText(word.word, { lang: "en-US", notify: true })}
              >
                <Volume2 size={13} />
              </button>
            </p>
          </div>
          <button className="details-button" type="button" onClick={() => onToggleExpanded(word)}>
            {expanded ? "Hide Details" : "View Details"}
            {expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          </button>
        </div>

        <p className="word-meaning">{word.meaning}</p>

        <div className="word-base-details">
          <section>
            <h3>EXAMPLE SENTENCES</h3>
            <blockquote>
              <p {...exampleSpeechHandlers}>"{word.example}"</p>
              <span>“{word.exampleTranslation}”</span>
            </blockquote>
          </section>
        </div>

        {expanded ? (
          <div className="word-details">
            {detailState?.loading ? <p className="word-detail-status">正在生成详情...</p> : null}
            {detailState?.error ? <p className="word-detail-status word-detail-error">{detailState.error}</p> : null}
            {detailState?.data ? <WordAiDetailsPanel details={detailState.data} word={word} /> : null}
          </div>
        ) : null}
      </div>

      <div className="mark-actions">
        <button type="button" disabled={Boolean(markingStatus)} onClick={() => requestMark("known")}>
          <CircleCheck size={17} />
          认识
        </button>
        <button
          type="button"
          disabled={Boolean(markingStatus)}
          onClick={() => requestMark("familiar")}
        >
          <CircleHelp size={17} />
          有印象
        </button>
        <button
          type="button"
          disabled={Boolean(markingStatus)}
          onClick={() => requestMark("unknown")}
        >
          <CircleX size={17} />
          不认识
        </button>
      </div>
    </article>
  );
}

function WordDetailLine({ line }: { line: string }) {
  const speechHandlers = useDelayedSpeech(line, 650, { lang: "en-US" });
  const canSpeak = /[A-Za-z]{2,}/.test(line);
  const isHeading = ["词根词源拆解", "核心释义", "搭配短句和实用例句", "衍生词和近义词辨析"].includes(line.trim());

  if (isHeading) return <h3>{line}</h3>;

  return (
    <p className={canSpeak ? "word-detail-readable-line" : undefined} {...(canSpeak ? speechHandlers : {})}>
      {line}
    </p>
  );
}

function WordAiDetailsPanel({ details, word }: { details: WordAiDetails; word: VocabularyWord }) {
  const synonymLine = word.synonyms.length ? `同义词：${word.synonyms.join(" / ")}` : "";
  const lines = [...details.content.split("\n").map((line) => line.trim()), synonymLine].filter(Boolean);

  return (
    <div className="word-detail-plain-text">
      {lines.map((line, index) => (
        <WordDetailLine key={`${line}-${index}`} line={line} />
      ))}
    </div>
  );
}

function StudyCompleteCard({
  count,
  onAddTenWords,
  onEnterQuiz,
}: {
  count: number;
  onAddTenWords: () => void;
  onEnterQuiz: () => void;
}) {
  return (
    <div className="study-complete-card">
      <div className="study-complete-icon">
        <Check size={34} />
      </div>
      <h2>本组学习已完成！</h2>
      <p>恭喜！你已经完成了本组 {count} 个单词的学习。</p>
      <div>
        <AppButton tone="primary" size="md" icon={<BrainCircuit size={20} />} onClick={onEnterQuiz}>
          立即测验
        </AppButton>
        <AppButton tone="outline" size="md" icon={<Repeat2 size={20} />} onClick={onAddTenWords}>
          再学 10 个
        </AppButton>
      </div>
    </div>
  );
}

function TranslationHeader({ onBack }: { onBack: () => void }) {
  return (
    <header className="translation-header">
      <div className="translation-header-inner">
        <div className="translation-brand">
          <button className="translation-back" type="button" onClick={onBack} aria-label="返回">
            <ArrowLeft size={16} />
          </button>
          <div className="wordmark" aria-label="Wordloop">
            Wordloop
          </div>
          <span className="translation-divider" />
          <span className="translation-title">快速测验</span>
        </div>
      </div>
    </header>
  );
}

function TranslationTestPage({
  onBack,
  onModeChange,
  onOpenSettings,
  onSubmit,
  onTranslationInputChange,
  prompt,
  settingsConfigured,
  status,
  translationError,
  translationInput,
  translationMode,
}: {
  onBack: () => void;
  onModeChange: (mode: TranslationMode) => void;
  onOpenSettings: () => void;
  onSubmit: () => void;
  prompt: TranslationPromptData | null;
  settingsConfigured: boolean;
  status: TranslationStatus;
  translationError: string;
  onTranslationInputChange: (value: string) => void;
  translationInput: string;
  translationMode: TranslationMode;
}) {
  const canSubmit = Boolean(prompt) && translationInput.trim().length > 0 && status !== "evaluating" && status !== "generating";
  const highlightedSource = prompt ? renderHighlightedText(prompt.source, prompt.highlightedWords) : null;
  const [listening, setListening] = useState(false);
  const [speechNotice, setSpeechNotice] = useState("");
  const recognitionRef = useRef<SpeechRecognitionLike | null>(null);

  useEffect(() => {
    return () => recognitionRef.current?.stop();
  }, []);

  function toggleSpeechInput() {
    if (listening) {
      recognitionRef.current?.stop();
      setListening(false);
      return;
    }

    const SpeechRecognitionConstructor = getSpeechRecognitionConstructor();
    if (!SpeechRecognitionConstructor) {
      setSpeechNotice("当前浏览器不支持语音输入，请使用 Chrome 或 Edge。");
      return;
    }

    const recognition = new SpeechRecognitionConstructor();
    recognition.continuous = false;
    recognition.interimResults = false;
    recognition.lang = translationMode === "enToZh" ? "zh-CN" : "en-US";
    recognition.onresult = (event) => {
      const transcript = Array.from(event.results)
        .map((result) => result[0]?.transcript ?? "")
        .join("")
        .trim();
      if (transcript) onTranslationInputChange(transcript);
      setSpeechNotice("");
    };
    recognition.onerror = () => {
      setSpeechNotice("语音输入失败，请检查麦克风权限后重试。");
      setListening(false);
    };
    recognition.onend = () => setListening(false);
    recognitionRef.current = recognition;
    setSpeechNotice("");
    setListening(true);
    recognition.start();
  }

  return (
    <main className="translation-page translation-input-page">
      <TranslationHeader onBack={onBack} />
      <section className="translation-input-main" aria-labelledby="translation-test-title">
        <div className="translation-card">
          <div className="translation-card-head">
            <h1 id="translation-test-title">段落翻译</h1>
            <div className="translation-mode-toggle" aria-label="翻译方向">
              <button
                className={translationMode === "enToZh" ? "active" : undefined}
                type="button"
                onClick={() => onModeChange("enToZh")}
              >
                英译中
              </button>
              <button
                className={translationMode === "zhToEn" ? "active" : undefined}
                type="button"
                onClick={() => onModeChange("zhToEn")}
              >
                中译英
              </button>
            </div>
          </div>

          <section className="ai-context-card" aria-label="AI 生成段落">
            <div className="ai-context-label">
              {status === "generating" ? <Loader2 className="loading-spinner" size={18} /> : <Sparkles size={18} />}
              <span>{status === "generating" ? "AI 正在生成段落中" : "本段落由ai根据所学词汇随机生成"}</span>
            </div>
            {status === "generating" ? (
              <div className="ai-context-skeleton" aria-label="段落生成中">
                <i />
                <i />
                <i />
                <i />
              </div>
            ) : (
              <p>{highlightedSource ?? "请先从学习页进入测验，或切换翻译方向重新生成段落。"}</p>
            )}
          </section>

          <section className="translation-input-block" aria-label="你的翻译">
            <label htmlFor="translation-answer">你的翻译</label>
            <div className="translation-textarea-wrap">
              <textarea
                id="translation-answer"
                value={translationInput}
                placeholder="请输入段落翻译"
                onChange={(event) => onTranslationInputChange(event.target.value)}
              />
              <button
                className={listening ? "speech-button speech-button-active" : "speech-button"}
                type="button"
                aria-label={listening ? "停止语音输入" : "开始语音输入"}
                title={listening ? "停止语音输入" : "开始语音输入"}
                onClick={toggleSpeechInput}
              >
                <Mic size={19} />
              </button>
            </div>
            {speechNotice ? <p className="speech-notice">{speechNotice}</p> : null}
          </section>

          <div className="translation-submit-row">
            <button className="translation-submit" disabled={!canSubmit} type="button" onClick={onSubmit}>
              {status === "evaluating" ? (
                <>
                  <Loader2 size={18} />
                  测评中
                </>
              ) : (
                "提交"
              )}
            </button>
            {!canSubmit ? <span className="translation-disabled-tip">Please type your answer first</span> : null}
          </div>
          {!settingsConfigured || translationError ? (
            <div className="translation-error">
              <Info size={16} />
              <span>{translationError || "使用 AI 测评前需要先配置 DeepSeek API Key。"}</span>
              {!settingsConfigured ? (
                <button type="button" onClick={onOpenSettings}>
                  去设置
                </button>
              ) : null}
            </div>
          ) : null}
        </div>
      </section>
    </main>
  );
}

function renderHighlightedText(source: string, words: string[]) {
  if (!words.length) return source;
  const escaped = words.map((word) => word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  const matcher = new RegExp(`(${escaped.join("|")})`, "gi");
  return source.split(matcher).map((part, index) =>
    words.some((word) => word.toLowerCase() === part.toLowerCase()) ? (
      <em key={`${part}-${index}`}>{part}</em>
    ) : (
      part
    ),
  );
}

function TranslationResultPage({
  onBack,
  onComplete,
  onRetry,
  result,
}: {
  onBack: () => void;
  onComplete: () => void;
  onRetry: () => void;
  result: TranslationEvaluationResult | null;
}) {
  if (!result) {
    return (
      <main className="translation-page translation-result-page">
        <TranslationHeader onBack={onBack} />
        <section className="translation-empty-result">
          <Info size={24} />
          <h1>暂无测评结果</h1>
          <p>请回到段落翻译页提交内容后再查看 AI 测评。</p>
          <button type="button" onClick={onRetry}>
            返回测验
          </button>
        </section>
      </main>
    );
  }

  return (
    <main className="translation-page translation-result-page">
      <TranslationHeader onBack={onBack} />
      <section className="translation-result-main" aria-label="AI 翻译测试结果">
        <div className="translation-score-section">
          <div className="translation-score-circle">
            <strong>{result.score}%</strong>
            <span>准确率</span>
          </div>
          <div className="translation-critique">
            <Sparkles size={22} />
            <div>
              <h1>AI 翻译表现点评</h1>
              <p>{result.critique}</p>
            </div>
          </div>
        </div>

        <section className="translation-comparison-card" aria-label="翻译对比分析">
          <div className="translation-original">
            <div className="mini-heading">
              <FileText size={11} />
              <span>原文</span>
            </div>
            <p>
              {result.source}
            </p>
          </div>

          <div className="translation-comparison-grid">
            <div>
              <div className="mini-heading mini-heading-danger">
                <UserRound size={10} />
                <span>您的翻译</span>
              </div>
              <p className="translation-user-copy">
                {result.userTranslation}
              </p>
            </div>
            <div>
              <div className="mini-heading">
                <CircleCheck size={12} />
                <span>正确参考答案</span>
              </div>
              <p className="translation-reference-copy">{result.reference}</p>
            </div>
          </div>
        </section>

        <section className="translation-review-section" aria-labelledby="translation-review-title">
          <div className="translation-review-head">
            <h2 id="translation-review-title">单词掌握度复盘</h2>
            <span aria-hidden="true" />
          </div>
          <div className="translation-review-table">
            <div className="translation-review-row translation-review-row-head">
              <span>单词</span>
              <span>中文释义</span>
              <span>掌握情况</span>
              <span>错误原因解析</span>
            </div>
            {result.reviewRows.map((row) => (
              <div className="translation-review-row" key={row.word}>
                <span>{row.word}</span>
                <span>{row.meaning}</span>
                <span className={`review-status review-status-${row.statusType}`}>
                  {row.statusType === "correct" ? <Check size={10} /> : <X size={10} />}
                  {row.status}
                </span>
                <span>{row.reason}</span>
              </div>
            ))}
          </div>
          <div className="translation-advice">
            <Info size={17} />
            <p>本次翻译出现了掌握情况欠佳的词汇。它们将在未来的学习中更频繁地出现以加强记忆。</p>
          </div>
        </section>

        <div className="translation-footer-actions">
          <button className="translation-retry" type="button" onClick={onRetry}>
            <RefreshCcw size={24} />
            再测一次
          </button>
          <button className="translation-complete" type="button" onClick={onComplete}>
            <CircleCheck size={24} />
            完成测验
          </button>
        </div>
      </section>
    </main>
  );
}

function MarkedSidebar({
  onUndoWord,
  session,
  words,
}: {
  onUndoWord: (wordId: string) => void;
  session: StudySession;
  words: VocabularyWord[];
}) {
  const groups: { status: MarkStatus; title: string }[] = [
    { status: "known", title: "认识" },
    { status: "familiar", title: "有印象" },
    { status: "unknown", title: "不认识" },
  ];

  return (
    <aside className="marked-sidebar">
      <div className="marked-card">
        <div className="marked-card-title">
          <BookmarkCheck size={22} />
          <h2>已标记单词</h2>
        </div>

        {groups.map((group) => {
          const markedWords = words.filter((word) => session.statuses[word.id] === group.status);
          return (
            <section className={`marked-group marked-group-${group.status}`} key={group.status}>
              <div className="marked-group-head">
                <span>
                  {group.title} ({markedWords.length})
                </span>
                <i />
              </div>
              <div className="marked-list">
                {markedWords.length ? (
                  markedWords.map((word) => (
                    <MarkedWordTag
                      key={word.id}
                      onUndoWord={onUndoWord}
                      status={group.status}
                      word={word}
                    />
                  ))
                ) : (
                  <p>暂无单词</p>
                )}
              </div>
            </section>
          );
        })}

        <div className="session-tip">
          <p>系统将会根据你标记的熟练程度安排复习频次</p>
        </div>
      </div>
    </aside>
  );
}

function MarkedWordTag({
  onUndoWord,
  status,
  word,
}: {
  onUndoWord: (wordId: string) => void;
  status: MarkStatus;
  word: VocabularyWord;
}) {
  const [removing, setRemoving] = useState(false);
  const removeTimeoutRef = useRef<number | null>(null);

  useEffect(() => {
    return () => {
      if (removeTimeoutRef.current) window.clearTimeout(removeTimeoutRef.current);
    };
  }, []);

  function requestUndo() {
    if (removing) return;
    setRemoving(true);
    removeTimeoutRef.current = window.setTimeout(() => {
      onUndoWord(word.id);
    }, 150);
  }

  return (
    <button
      className={removing ? "marked-word-tag marked-word-tag-removing" : "marked-word-tag"}
      data-status={status}
      disabled={removing}
      type="button"
      onClick={requestUndo}
    >
      {word.word}
      <X size={13} />
    </button>
  );
}

function ReviewCard({ data, onStartReview }: { data: ReviewData; onStartReview: () => void }) {
  const done = data.state === "todayDone" || data.state === "allDone";
  const hasCta = data.state !== "allDone";
  const statusLabel =
    data.state === "notStarted" ? "尚未开始复习" : done ? "完成度已达标" : data.state === "inProgress" ? "复习中" : data.status;

  return (
    <article className={`review-card review-card-${data.state}`}>
      <div className="review-card-head">
        <div>
          <h3>今日复习进度</h3>
          <p>{statusLabel}</p>
        </div>
        <ProgressRing percent={data.percent} />
      </div>

      {done ? (
        <CompletionMessage
          title={
            data.state === "todayDone"
              ? "太棒了！ 今日复习任务已全部达成！"
              : "太棒了！ 复习任务已全部达成！"
          }
        />
      ) : null}

      <div className="review-metrics">
        <MetricCard icon={<CalendarClock size={16} />} label="待复习" value={data.due} />
        <MetricCard
          danger={data.overdue > 0}
          icon={<Hourglass size={16} />}
          label="已过期"
          value={data.overdue}
        />
      </div>

      {hasCta ? (
        <AppButton tone="primary" fullWidth onClick={onStartReview}>
          {done ? "继续复习" : "立即复习"}
        </AppButton>
      ) : null}
    </article>
  );
}

function ProgressRing({ percent }: { percent: number }) {
  const stroke = 4;
  const radius = 22;
  const circumference = 2 * Math.PI * radius;
  const offset = circumference - (percent / 100) * circumference;

  return (
    <div className="progress-ring" aria-label={`复习进度 ${percent}%`}>
      <svg viewBox="0 0 48 48" aria-hidden="true">
        <circle cx="24" cy="24" r={radius} strokeWidth={stroke} className="ring-track" />
        <circle
          cx="24"
          cy="24"
          r={radius}
          strokeWidth={stroke}
          className="ring-value"
          strokeDasharray={circumference}
          strokeDashoffset={offset}
        />
      </svg>
      <span>{percent}%</span>
    </div>
  );
}

function CompletionMessage({ title }: { title: string }) {
  return (
    <div className="completion-message">
      <div className="completion-icon">
        <Check size={46} strokeWidth={2.4} />
      </div>
      <p>{title}</p>
    </div>
  );
}

function MetricCard({
  danger,
  icon,
  label,
  suffix = "",
  value,
}: {
  danger?: boolean;
  icon: ReactNode;
  label: string;
  suffix?: string;
  value: number;
}) {
  return (
    <div className={danger ? "metric-card metric-card-danger" : "metric-card"}>
      <div className="metric-icon">{icon}</div>
      <div>
        <span>{label}</span>
        <strong>
          {value}
          {suffix ? <small>{suffix}</small> : null}
        </strong>
      </div>
    </div>
  );
}

function LearningRecordCard({ stats }: { stats: RecordStat[] }) {
  const learnedCount = Number(stats.find((stat) => stat.label === "已学习")?.value ?? 0);
  const realCalendarDays = getLearningCalendarDays(learnedCount > 0);

  return (
    <section className="record-card" aria-labelledby="record-title">
      <div className="record-head">
        <h3 id="record-title">学习记录</h3>
        <span>2026年7月</span>
      </div>

      <div className="week-row" aria-hidden="true">
        {["一", "二", "三", "四", "五", "六", "日"].map((day) => (
          <span key={day}>{day}</span>
        ))}
      </div>

      <div className="calendar-grid">
        {realCalendarDays.map((day, index) => (
          <div key={`${day.label}-${index}`} className={`calendar-day calendar-${day.status}`}>
            <span>{day.label}</span>
            {day.status === "done" || day.status === "missed" ? <i /> : null}
          </div>
        ))}
      </div>

      <div className="record-stats">
        {stats.map((stat) => (
          <div key={stat.label}>
            <span>{stat.label}</span>
            <strong>
              {stat.value}
              <small>{stat.suffix}</small>
            </strong>
          </div>
        ))}
      </div>
    </section>
  );
}

function AppButton({
  children,
  disabled = false,
  fullWidth,
  icon,
  onClick,
  size = "sm",
  tone,
}: {
  children: ReactNode;
  disabled?: boolean;
  fullWidth?: boolean;
  icon?: ReactNode;
  onClick?: () => void;
  size?: ButtonSize;
  tone: ButtonTone;
}) {
  return (
    <button
      className={`app-button app-button-${tone} app-button-${size}${fullWidth ? " app-button-full" : ""}`}
      disabled={disabled}
      onClick={onClick}
      type="button"
    >
      {icon}
      <span>{children}</span>
    </button>
  );
}

export default App;
