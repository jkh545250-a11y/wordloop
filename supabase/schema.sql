-- Wordloop Supabase schema
-- Run this in Supabase Dashboard -> SQL Editor -> New query.

create extension if not exists pgcrypto;

create table if not exists public.decks (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  local_id text not null,
  name text not null,
  description text not null default '',
  is_builtin boolean not null default false,
  is_custom boolean not null default true,
  ai_generated boolean not null default false,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, local_id)
);

create table if not exists public.words (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  deck_id uuid not null references public.decks(id) on delete cascade,
  local_id text not null,
  word text not null,
  phonetic text not null default '',
  meaning text not null default '',
  example text not null default '',
  example_translation text not null default '',
  synonyms jsonb not null default '[]'::jsonb,
  extra_meanings jsonb not null default '[]'::jsonb,
  extra_examples jsonb not null default '[]'::jsonb,
  sort_order integer not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (deck_id, local_id)
);

create table if not exists public.word_progress (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  deck_id uuid not null references public.decks(id) on delete cascade,
  word_id uuid not null references public.words(id) on delete cascade,
  status text not null check (status in ('known', 'familiar', 'unknown')),
  learned_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, word_id)
);

create table if not exists public.review_cycles (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  deck_id uuid not null references public.decks(id) on delete cascade,
  word_id uuid not null references public.words(id) on delete cascade,
  cycles integer not null default 0 check (cycles >= 0),
  mastered boolean not null default false,
  updated_at timestamptz not null default now(),
  unique (user_id, word_id)
);

create table if not exists public.review_schedule (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  deck_id uuid not null references public.decks(id) on delete cascade,
  word_id uuid not null references public.words(id) on delete cascade,
  stage integer not null default 0 check (stage >= 0),
  ease numeric not null default 1,
  due_at timestamptz not null,
  last_reviewed_at timestamptz,
  lapses integer not null default 0 check (lapses >= 0),
  correct_streak integer not null default 0 check (correct_streak >= 0),
  source text not null default 'study' check (source in ('study', 'quiz', 'review')),
  last_rating text not null default 'legacy',
  updated_at timestamptz not null default now(),
  unique (user_id, word_id)
);

create table if not exists public.daily_review_stats (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  review_date date not null,
  reviewed_count integer not null default 0 check (reviewed_count >= 0),
  updated_at timestamptz not null default now(),
  unique (user_id, review_date)
);

create table if not exists public.daily_reviewed_words (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  review_date date not null,
  word_ids jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now(),
  unique (user_id, review_date)
);

create table if not exists public.deck_overrides (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  local_id text not null,
  name text,
  deleted boolean not null default false,
  updated_at timestamptz not null default now(),
  unique (user_id, local_id)
);

create table if not exists public.app_settings (
  user_id uuid primary key references auth.users(id) on delete cascade,
  daily_study_count integer not null default 20 check (daily_study_count between 1 and 100),
  daily_review_goal integer not null default 30 check (daily_review_goal between 1 and 200),
  updated_at timestamptz not null default now()
);

create index if not exists decks_user_id_idx on public.decks(user_id);
create index if not exists words_deck_id_idx on public.words(deck_id);
create index if not exists word_progress_user_id_idx on public.word_progress(user_id);
create index if not exists review_cycles_user_id_idx on public.review_cycles(user_id);
create index if not exists review_schedule_user_id_idx on public.review_schedule(user_id);
create index if not exists review_schedule_due_at_idx on public.review_schedule(due_at);
create index if not exists daily_review_stats_user_id_idx on public.daily_review_stats(user_id);
create index if not exists daily_reviewed_words_user_id_idx on public.daily_reviewed_words(user_id);
create index if not exists deck_overrides_user_id_idx on public.deck_overrides(user_id);

alter table public.decks enable row level security;
alter table public.words enable row level security;
alter table public.word_progress enable row level security;
alter table public.review_cycles enable row level security;
alter table public.review_schedule enable row level security;
alter table public.daily_review_stats enable row level security;
alter table public.daily_reviewed_words enable row level security;
alter table public.deck_overrides enable row level security;
alter table public.app_settings enable row level security;

drop policy if exists "Users can manage own decks" on public.decks;
create policy "Users can manage own decks"
on public.decks
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own words" on public.words;
create policy "Users can manage own words"
on public.words
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own word progress" on public.word_progress;
create policy "Users can manage own word progress"
on public.word_progress
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own review cycles" on public.review_cycles;
create policy "Users can manage own review cycles"
on public.review_cycles
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own review schedule" on public.review_schedule;
create policy "Users can manage own review schedule"
on public.review_schedule
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own daily review stats" on public.daily_review_stats;
create policy "Users can manage own daily review stats"
on public.daily_review_stats
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own daily reviewed words" on public.daily_reviewed_words;
create policy "Users can manage own daily reviewed words"
on public.daily_reviewed_words
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own deck overrides" on public.deck_overrides;
create policy "Users can manage own deck overrides"
on public.deck_overrides
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

drop policy if exists "Users can manage own app settings" on public.app_settings;
create policy "Users can manage own app settings"
on public.app_settings
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());
