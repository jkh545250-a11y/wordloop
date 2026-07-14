-- Wordloop Supabase grants
-- Run this after schema.sql if the app gets 403 Forbidden when reading/writing tables.

grant usage on schema public to anon, authenticated;

grant all privileges on table public.decks to authenticated;
grant all privileges on table public.words to authenticated;
grant all privileges on table public.word_progress to authenticated;
grant all privileges on table public.review_cycles to authenticated;
grant all privileges on table public.review_schedule to authenticated;
grant all privileges on table public.daily_review_stats to authenticated;
grant all privileges on table public.daily_reviewed_words to authenticated;
grant all privileges on table public.deck_overrides to authenticated;
grant all privileges on table public.app_settings to authenticated;

grant usage, select on all sequences in schema public to authenticated;

alter default privileges in schema public grant all privileges on tables to authenticated;
alter default privileges in schema public grant usage, select on sequences to authenticated;
