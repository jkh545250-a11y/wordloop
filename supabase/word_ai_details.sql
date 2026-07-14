-- Wordloop AI word detail cache
-- Run this in Supabase SQL Editor before using AI-generated word details.

create table if not exists public.word_ai_details (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  word_key text not null,
  deck_local_id text not null,
  word_local_id text not null,
  details jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, word_key)
);

create index if not exists word_ai_details_user_id_idx on public.word_ai_details(user_id);

alter table public.word_ai_details enable row level security;

drop policy if exists "Users can manage own word ai details" on public.word_ai_details;
create policy "Users can manage own word ai details"
on public.word_ai_details
for all
to authenticated
using (user_id = auth.uid())
with check (user_id = auth.uid());

grant all privileges on table public.word_ai_details to authenticated;
