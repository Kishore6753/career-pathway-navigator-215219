-- Career Navigator MVP - Supabase Schema and RLS disablement
-- This script creates core tables and disables RLS temporarily for MVP seeding.

-- Extensions
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- Roles
CREATE TABLE IF NOT EXISTS public.roles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text NOT NULL UNIQUE,
  name text NOT NULL,
  description text,
  created_at timestamptz DEFAULT now()
);

-- Competencies
CREATE TABLE IF NOT EXISTS public.competencies (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text NOT NULL UNIQUE,
  name text NOT NULL,
  category text,
  description text,
  created_at timestamptz DEFAULT now()
);

-- Mapping: Role -> Competencies with target level
CREATE TABLE IF NOT EXISTS public.role_competencies (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  competency_id uuid NOT NULL REFERENCES public.competencies(id) ON DELETE CASCADE,
  target_level integer NOT NULL,
  created_at timestamptz DEFAULT now(),
  UNIQUE (role_id, competency_id)
);

-- Users (MVP: basic profile only)
CREATE TABLE IF NOT EXISTS public.user_profiles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid, -- optional linkage to auth.users
  full_name text,
  email text UNIQUE,
  current_role_id uuid REFERENCES public.roles(id),
  created_at timestamptz DEFAULT now()
);

-- Mapping: User -> Competencies with level
CREATE TABLE IF NOT EXISTS public.user_competencies (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_profile_id uuid NOT NULL REFERENCES public.user_profiles(id) ON DELETE CASCADE,
  competency_id uuid NOT NULL REFERENCES public.competencies(id) ON DELETE CASCADE,
  level integer,
  created_at timestamptz DEFAULT now(),
  UNIQUE (user_profile_id, competency_id)
);

-- Role adjacency (similarity/transition score)
CREATE TABLE IF NOT EXISTS public.role_adjacency (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  source_role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  target_role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  score numeric,
  created_at timestamptz DEFAULT now(),
  UNIQUE (source_role_id, target_role_id)
);

-- Learning resources linked to competencies
CREATE TABLE IF NOT EXISTS public.learning_resources (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  competency_id uuid REFERENCES public.competencies(id) ON DELETE CASCADE,
  title text NOT NULL,
  url text NOT NULL,
  provider text,
  created_at timestamptz DEFAULT now(),
  UNIQUE (competency_id, url)
);

-- Disable RLS for MVP seeding (re-enable with ENABLE ROW LEVEL SECURITY when ready)
ALTER TABLE public.roles                DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.competencies         DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.role_competencies    DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_profiles        DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_competencies    DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.role_adjacency       DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.learning_resources   DISABLE ROW LEVEL SECURITY;
