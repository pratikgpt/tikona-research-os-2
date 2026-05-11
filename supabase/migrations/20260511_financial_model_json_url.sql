alter table public.research_sessions
add column if not exists financial_model_json_url text;
