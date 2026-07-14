import { builtinDecks, type VocabularyWord } from "../data/builtinDecks";
import { supabase } from "./supabase";

type LocalDeck = {
  id: string;
  name: string;
  description: string;
  words: VocabularyWord[];
  aiGenerated?: boolean;
  custom?: boolean;
};

export type LocalProgress = Record<string, Record<string, "known" | "familiar" | "unknown">>;
export type LocalMastery = Record<string, Record<string, true>>;
export type LocalReviewCycles = Record<string, Record<string, number>>;
export type LocalReviewStats = Record<string, number>;
export type LocalDailyReviewedWords = Record<string, string[]>;
export type LocalDeckOverrides = Record<string, { deleted?: boolean; name?: string }>;
export type LocalReviewSchedule = Record<
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
      source: "study" | "quiz" | "review";
      lastRating: string;
    }
  >
>;

export type LocalSettings = {
  dailyStudyCount?: number;
  dailyReviewGoal?: number;
};

export type SupabaseHydrationData = {
  userDecks: LocalDeck[];
  progress: LocalProgress;
  mastery: LocalMastery;
  reviewCycles: LocalReviewCycles;
  reviewStats: LocalReviewStats;
  dailyReviewedWords: LocalDailyReviewedWords;
  deckOverrides: LocalDeckOverrides;
  reviewSchedule: LocalReviewSchedule;
  settings: LocalSettings;
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
const masteryReviewCycleTarget = 5;

function getScopedStorageKey(key: string, userId: string) {
  return `${key}.${userId}`;
}

function readStorageRaw(key: string, userId: string) {
  if (typeof window === "undefined") return null;
  return window.localStorage.getItem(getScopedStorageKey(key, userId));
}

function readStorage<T>(key: string, fallback: T, userId: string): T {
  if (typeof window === "undefined") return fallback;
  try {
    const raw = readStorageRaw(key, userId);
    return raw ? { ...fallback, ...JSON.parse(raw) } : fallback;
  } catch {
    return fallback;
  }
}

function readStorageArray<T>(key: string, userId: string): T[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = readStorageRaw(key, userId);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function readLocalData(userId: string) {
  const deckOverrides = readStorage<LocalDeckOverrides>(deckOverridesStorageKey, {}, userId);
  const activeBuiltinDecks = builtinDecks
    .filter((deck) => !deckOverrides[deck.id]?.deleted)
    .map((deck) => ({ ...deck, name: deckOverrides[deck.id]?.name ?? deck.name }));
  return {
    decks: [...activeBuiltinDecks, ...readStorageArray<LocalDeck>(userDecksStorageKey, userId)] as LocalDeck[],
    deckOverrides,
    mastery: readStorage<LocalMastery>(masteryStorageKey, {}, userId),
    progress: readStorage<LocalProgress>(progressStorageKey, {}, userId),
    reviewCycles: readStorage<LocalReviewCycles>(reviewCyclesStorageKey, {}, userId),
    reviewStats: readStorage<LocalReviewStats>(reviewStatsStorageKey, {}, userId),
    dailyReviewedWords: readStorage<LocalDailyReviewedWords>(dailyReviewedWordsStorageKey, {}, userId),
    reviewSchedule: readStorage<LocalReviewSchedule>(reviewScheduleStorageKey, {}, userId),
    settings: readStorage<LocalSettings>(settingsStorageKey, {}, userId),
  };
}

function isMastered(deckId: string, wordId: string, mastery: LocalMastery, reviewCycles: LocalReviewCycles) {
  return Boolean(mastery[deckId]?.[wordId] && (reviewCycles[deckId]?.[wordId] ?? 0) >= masteryReviewCycleTarget);
}

function isOptionalTableError(error: unknown) {
  const candidate = error as { code?: string; message?: string } | null;
  const message = candidate?.message?.toLowerCase() ?? "";
  return (
    candidate?.code === "42P01" ||
    candidate?.code === "PGRST205" ||
    candidate?.code === "42501" ||
    message.includes("does not exist") ||
    message.includes("could not find the table") ||
    message.includes("permission denied")
  );
}

async function throwUnlessOptionalTableError(error: unknown) {
  if (!error || isOptionalTableError(error)) return;
  throw error;
}

export async function syncLocalDataToSupabase(userId: string) {
  const { decks, deckOverrides, mastery, progress, reviewCycles, reviewStats, dailyReviewedWords, reviewSchedule, settings } =
    readLocalData(userId);
  if (!decks.length) return;

  const deletedBuiltinIds = Object.entries(deckOverrides)
    .filter(([, override]) => override.deleted)
    .map(([deckId]) => deckId);
  if (deletedBuiltinIds.length) {
    const { error } = await supabase
      .from("decks")
      .delete()
      .eq("user_id", userId)
      .eq("is_builtin", true)
      .in("local_id", deletedBuiltinIds);
    if (error) throw error;
  }

  const customDeckIds = decks.filter((deck) => deck.custom).map((deck) => deck.id);
  let staleCustomDeckDelete = supabase
    .from("decks")
    .delete()
    .eq("user_id", userId)
    .eq("is_custom", true);

  if (customDeckIds.length) {
    staleCustomDeckDelete = staleCustomDeckDelete.not("local_id", "in", `(${customDeckIds.map((id) => `"${id}"`).join(",")})`);
  }

  const { error: staleDecksError } = await staleCustomDeckDelete;
  if (staleDecksError) throw staleDecksError;

  const requiredCleanupResults = await Promise.all([
    supabase.from("word_progress").delete().eq("user_id", userId),
    supabase.from("review_cycles").delete().eq("user_id", userId),
    supabase.from("daily_review_stats").delete().eq("user_id", userId),
  ]);
  const failedRequiredCleanup = requiredCleanupResults.find((result) => result.error);
  if (failedRequiredCleanup?.error) throw failedRequiredCleanup.error;

  const optionalCleanupResults = await Promise.all([
    supabase.from("review_schedule").delete().eq("user_id", userId),
    supabase.from("daily_reviewed_words").delete().eq("user_id", userId),
    supabase.from("deck_overrides").delete().eq("user_id", userId),
  ]);
  await Promise.all(optionalCleanupResults.map((result) => throwUnlessOptionalTableError(result.error)));

  const { data: syncedDecks, error: decksError } = await supabase
    .from("decks")
    .upsert(
      decks.map((deck) => ({
        user_id: userId,
        local_id: deck.id,
        name: deck.name,
        description: deck.description ?? "",
        is_builtin: !deck.custom,
        is_custom: Boolean(deck.custom),
        ai_generated: Boolean(deck.aiGenerated),
      })),
      { onConflict: "user_id,local_id" },
    )
    .select("id, local_id");

  if (decksError) throw decksError;

  const deckIdByLocalId = new Map((syncedDecks ?? []).map((deck) => [deck.local_id as string, deck.id as string]));
  const wordRows = decks.flatMap((deck) => {
    const deckUuid = deckIdByLocalId.get(deck.id);
    if (!deckUuid) return [];
    return deck.words.map((word, index) => ({
      user_id: userId,
      deck_id: deckUuid,
      local_id: word.id,
      word: word.word,
      phonetic: word.phonetic,
      meaning: word.meaning,
      example: word.example,
      example_translation: word.exampleTranslation,
      synonyms: word.synonyms,
      extra_meanings: word.extraMeanings,
      extra_examples: word.extraExamples,
      sort_order: index,
    }));
  });

  let wordIdByDeckAndLocalId = new Map<string, string>();
  if (wordRows.length) {
    const { data: syncedWords, error: wordsError } = await supabase
      .from("words")
      .upsert(wordRows, { onConflict: "deck_id,local_id" })
      .select("id, deck_id, local_id");

    if (wordsError) throw wordsError;
    wordIdByDeckAndLocalId = new Map(
      (syncedWords ?? []).map((word) => [`${word.deck_id}:${word.local_id}`, word.id as string]),
    );
  }

  const progressRows = Object.entries(progress).flatMap(([deckLocalId, deckProgress]) => {
    const deckUuid = deckIdByLocalId.get(deckLocalId);
    if (!deckUuid) return [];
    return Object.entries(deckProgress).flatMap(([wordLocalId, status]) => {
      const wordUuid = wordIdByDeckAndLocalId.get(`${deckUuid}:${wordLocalId}`);
      if (!wordUuid) return [];
      return [
        {
          user_id: userId,
          deck_id: deckUuid,
          word_id: wordUuid,
          status,
        },
      ];
    });
  });

  if (progressRows.length) {
    const { error } = await supabase.from("word_progress").upsert(progressRows, { onConflict: "user_id,word_id" });
    if (error) throw error;
  }

  const reviewCycleRows = Object.entries(reviewCycles).flatMap(([deckLocalId, deckCycles]) => {
    const deckUuid = deckIdByLocalId.get(deckLocalId);
    if (!deckUuid) return [];
    return Object.entries(deckCycles).flatMap(([wordLocalId, cycles]) => {
      const wordUuid = wordIdByDeckAndLocalId.get(`${deckUuid}:${wordLocalId}`);
      if (!wordUuid) return [];
      return [
        {
          user_id: userId,
          deck_id: deckUuid,
          word_id: wordUuid,
          cycles,
          mastered: isMastered(deckLocalId, wordLocalId, mastery, reviewCycles),
        },
      ];
    });
  });

  if (reviewCycleRows.length) {
    const { error } = await supabase.from("review_cycles").upsert(reviewCycleRows, { onConflict: "user_id,word_id" });
    if (error) throw error;
  }

  const reviewScheduleRows = Object.entries(reviewSchedule).flatMap(([deckLocalId, deckSchedule]) => {
    const deckUuid = deckIdByLocalId.get(deckLocalId);
    if (!deckUuid) return [];
    return Object.entries(deckSchedule).flatMap(([wordLocalId, memory]) => {
      const wordUuid = wordIdByDeckAndLocalId.get(`${deckUuid}:${wordLocalId}`);
      if (!wordUuid) return [];
      return [
        {
          user_id: userId,
          deck_id: deckUuid,
          word_id: wordUuid,
          stage: memory.stage,
          ease: memory.ease,
          due_at: memory.dueAt,
          last_reviewed_at: memory.lastReviewedAt ?? null,
          lapses: memory.lapses,
          correct_streak: memory.correctStreak,
          source: memory.source,
          last_rating: memory.lastRating,
        },
      ];
    });
  });

  if (reviewScheduleRows.length) {
    const { error } = await supabase.from("review_schedule").upsert(reviewScheduleRows, { onConflict: "user_id,word_id" });
    await throwUnlessOptionalTableError(error);
  }

  const reviewStatRows = Object.entries(reviewStats).map(([reviewDate, reviewedCount]) => ({
    user_id: userId,
    review_date: reviewDate,
    reviewed_count: reviewedCount,
  }));

  if (reviewStatRows.length) {
    const { error } = await supabase
      .from("daily_review_stats")
      .upsert(reviewStatRows, { onConflict: "user_id,review_date" });
    if (error) throw error;
  }

  const dailyReviewedRows = Object.entries(dailyReviewedWords).map(([reviewDate, wordIds]) => ({
    user_id: userId,
    review_date: reviewDate,
    word_ids: wordIds,
  }));

  if (dailyReviewedRows.length) {
    const { error } = await supabase
      .from("daily_reviewed_words")
      .upsert(dailyReviewedRows, { onConflict: "user_id,review_date" });
    await throwUnlessOptionalTableError(error);
  }

  const deckOverrideRows = Object.entries(deckOverrides).map(([localId, override]) => ({
    user_id: userId,
    local_id: localId,
    name: override.name ?? null,
    deleted: Boolean(override.deleted),
  }));

  if (deckOverrideRows.length) {
    const { error } = await supabase.from("deck_overrides").upsert(deckOverrideRows, { onConflict: "user_id,local_id" });
    await throwUnlessOptionalTableError(error);
  }

  const { error: settingsError } = await supabase.from("app_settings").upsert(
    {
      user_id: userId,
      daily_study_count: settings.dailyStudyCount ?? 20,
      daily_review_goal: settings.dailyReviewGoal ?? 30,
    },
    { onConflict: "user_id" },
  );

  if (settingsError) throw settingsError;
}

export async function loadSupabaseData(userId: string): Promise<SupabaseHydrationData | null> {
  const { data: decks, error: decksError } = await supabase
    .from("decks")
    .select("id, local_id, name, description, is_custom, ai_generated")
    .eq("user_id", userId);

  if (decksError) throw decksError;
  if (!decks?.length) return null;

  const deckUuidByLocalId = new Map(decks.map((deck) => [deck.local_id as string, deck.id as string]));
  const deckLocalIdByUuid = new Map(decks.map((deck) => [deck.id as string, deck.local_id as string]));

  const { data: words, error: wordsError } = await supabase
    .from("words")
    .select(
      "id, deck_id, local_id, word, phonetic, meaning, example, example_translation, synonyms, extra_meanings, extra_examples, sort_order",
    )
    .eq("user_id", userId)
    .order("sort_order");

  if (wordsError) throw wordsError;

  const wordLocalByUuid = new Map<string, { deckLocalId: string; wordLocalId: string }>();
  const wordsByDeckUuid = new Map<string, VocabularyWord[]>();

  (words ?? []).forEach((word) => {
    const deckLocalId = deckLocalIdByUuid.get(word.deck_id as string);
    if (!deckLocalId) return;
    wordLocalByUuid.set(word.id as string, { deckLocalId, wordLocalId: word.local_id as string });
    const normalizedWord: VocabularyWord = {
      id: word.local_id as string,
      word: word.word as string,
      phonetic: (word.phonetic as string) ?? "",
      meaning: (word.meaning as string) ?? "",
      example: (word.example as string) ?? "",
      exampleTranslation: (word.example_translation as string) ?? "",
      synonyms: Array.isArray(word.synonyms) ? word.synonyms : [],
      extraMeanings: Array.isArray(word.extra_meanings) ? word.extra_meanings : [],
      extraExamples: Array.isArray(word.extra_examples) ? word.extra_examples : [],
    };
    wordsByDeckUuid.set(word.deck_id as string, [...(wordsByDeckUuid.get(word.deck_id as string) ?? []), normalizedWord]);
  });

  const userDecks: LocalDeck[] = decks
    .filter((deck) => Boolean(deck.is_custom))
    .map((deck) => ({
      id: deck.local_id as string,
      name: deck.name as string,
      description: (deck.description as string) ?? "",
      words: wordsByDeckUuid.get(deck.id as string) ?? [],
      aiGenerated: Boolean(deck.ai_generated),
      custom: true,
    }));

  const progress: LocalProgress = {};
  const { data: progressRows, error: progressError } = await supabase
    .from("word_progress")
    .select("deck_id, word_id, status")
    .eq("user_id", userId);

  if (progressError) throw progressError;

  (progressRows ?? []).forEach((row) => {
    const deckLocalId = deckLocalIdByUuid.get(row.deck_id as string);
    const wordLocal = wordLocalByUuid.get(row.word_id as string);
    if (!deckLocalId || !wordLocal) return;
    progress[deckLocalId] = { ...progress[deckLocalId], [wordLocal.wordLocalId]: row.status as "known" | "familiar" | "unknown" };
  });

  const mastery: LocalMastery = {};
  const reviewCycles: LocalReviewCycles = {};
  const reviewSchedule: LocalReviewSchedule = {};
  const { data: reviewCycleRows, error: reviewCyclesError } = await supabase
    .from("review_cycles")
    .select("deck_id, word_id, cycles, mastered")
    .eq("user_id", userId);

  if (reviewCyclesError) throw reviewCyclesError;

  (reviewCycleRows ?? []).forEach((row) => {
    const deckLocalId = deckLocalIdByUuid.get(row.deck_id as string);
    const wordLocal = wordLocalByUuid.get(row.word_id as string);
    if (!deckLocalId || !wordLocal) return;
    const cycles = Number(row.cycles ?? 0);
    reviewCycles[deckLocalId] = { ...reviewCycles[deckLocalId], [wordLocal.wordLocalId]: cycles };
    if (row.mastered && cycles >= masteryReviewCycleTarget) {
      mastery[deckLocalId] = { ...mastery[deckLocalId], [wordLocal.wordLocalId]: true };
    }
  });

  const { data: reviewScheduleRows, error: reviewScheduleError } = await supabase
    .from("review_schedule")
    .select("deck_id, word_id, stage, ease, due_at, last_reviewed_at, lapses, correct_streak, source, last_rating")
    .eq("user_id", userId);

  await throwUnlessOptionalTableError(reviewScheduleError);

  (reviewScheduleRows ?? []).forEach((row) => {
    const deckLocalId = deckLocalIdByUuid.get(row.deck_id as string);
    const wordLocal = wordLocalByUuid.get(row.word_id as string);
    if (!deckLocalId || !wordLocal) return;
    reviewSchedule[deckLocalId] = {
      ...reviewSchedule[deckLocalId],
      [wordLocal.wordLocalId]: {
        stage: Number(row.stage ?? 0),
        ease: Number(row.ease ?? 1),
        dueAt: (row.due_at as string) ?? new Date(0).toISOString(),
        lastReviewedAt: (row.last_reviewed_at as string | null) ?? undefined,
        lapses: Number(row.lapses ?? 0),
        correctStreak: Number(row.correct_streak ?? 0),
        source: (row.source as "study" | "quiz" | "review") ?? "study",
        lastRating: (row.last_rating as string) ?? "legacy",
      },
    };
  });

  const reviewStats: LocalReviewStats = {};
  const { data: reviewStatRows, error: reviewStatsError } = await supabase
    .from("daily_review_stats")
    .select("review_date, reviewed_count")
    .eq("user_id", userId);

  if (reviewStatsError) throw reviewStatsError;

  (reviewStatRows ?? []).forEach((row) => {
    reviewStats[row.review_date as string] = Number(row.reviewed_count ?? 0);
  });

  const dailyReviewedWords: LocalDailyReviewedWords = {};
  const { data: dailyReviewedRows, error: dailyReviewedError } = await supabase
    .from("daily_reviewed_words")
    .select("review_date, word_ids")
    .eq("user_id", userId);

  await throwUnlessOptionalTableError(dailyReviewedError);

  (dailyReviewedRows ?? []).forEach((row) => {
    dailyReviewedWords[row.review_date as string] = Array.isArray(row.word_ids) ? (row.word_ids as string[]) : [];
  });

  const deckOverrides: LocalDeckOverrides = {};
  const { data: deckOverrideRows, error: deckOverridesError } = await supabase
    .from("deck_overrides")
    .select("local_id, name, deleted")
    .eq("user_id", userId);

  await throwUnlessOptionalTableError(deckOverridesError);

  (deckOverrideRows ?? []).forEach((row) => {
    deckOverrides[row.local_id as string] = {
      name: (row.name as string | null) ?? undefined,
      deleted: Boolean(row.deleted),
    };
  });

  const { data: settingsRow, error: settingsError } = await supabase
    .from("app_settings")
    .select("daily_study_count, daily_review_goal")
    .eq("user_id", userId)
    .maybeSingle();

  if (settingsError) throw settingsError;

  return {
    userDecks,
    progress,
    mastery,
    reviewCycles,
    reviewStats,
    dailyReviewedWords,
    deckOverrides,
    reviewSchedule,
    settings: settingsRow
      ? {
          dailyStudyCount: Number(settingsRow.daily_study_count ?? 20),
          dailyReviewGoal: Number(settingsRow.daily_review_goal ?? 30),
        }
      : {},
  };
}
