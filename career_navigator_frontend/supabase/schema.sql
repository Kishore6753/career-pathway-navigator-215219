-- Career Navigator MVP - Supabase Schema (Idempotent) and RLS disablement
-- This script creates/ensures core tables and disables RLS for MVP seeding.
-- Safe to run multiple times.

-- Extensions
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- =====================================================================================
-- USERS (App users with role-based access)
-- Note: Supabase auth users live in auth.users; we optionally link via auth_user_id
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.users (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  auth_user_id uuid UNIQUE,
  email text UNIQUE,
  app_role text NOT NULL DEFAULT 'user',
  created_at timestamptz DEFAULT now()
);

-- Add FK to auth.users if not present
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM pg_constraint
    WHERE conname = 'users_auth_user_fk'
  ) THEN
    ALTER TABLE public.users
      ADD CONSTRAINT users_auth_user_fk
      FOREIGN KEY (auth_user_id) REFERENCES auth.users(id) ON DELETE SET NULL;
  END IF;
END$$;

-- Enforce role check if not present
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM pg_constraint
    WHERE conname = 'users_app_role_chk'
  ) THEN
    ALTER TABLE public.users
      ADD CONSTRAINT users_app_role_chk
      CHECK (app_role IN ('user','admin'));
  END IF;
END$$;

-- =====================================================================================
-- ROLES
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.roles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text NOT NULL UNIQUE,
  name text NOT NULL,
  description text,
  created_at timestamptz DEFAULT now()
);

-- Helpful index for name-based lookups
CREATE INDEX IF NOT EXISTS idx_roles_name ON public.roles (name);

-- =====================================================================================
-- ROLE DESCRIPTIONS (for multiple sources/versions per role)
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.role_descriptions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  source text DEFAULT 'card',        -- e.g., 'card', 'manual', 'import'
  summary text,                      -- optional short summary/abstract
  content text NOT NULL,             -- full body/text
  created_at timestamptz DEFAULT now(),
  UNIQUE (role_id, source)
);

CREATE INDEX IF NOT EXISTS idx_role_descriptions_role ON public.role_descriptions (role_id);

-- =====================================================================================
-- COMPETENCIES
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.competencies (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text NOT NULL UNIQUE,
  name text NOT NULL,
  category text,
  description text,
  created_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_competencies_name ON public.competencies (name);
CREATE INDEX IF NOT EXISTS idx_competencies_category ON public.competencies (category);

-- =====================================================================================
-- ROLE → COMPETENCIES with target level and weight
-- Levels map (typical): 1=Beginner, 2=Intermediate, 3=Advanced, 4=Master
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.role_competencies (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  competency_id uuid NOT NULL REFERENCES public.competencies(id) ON DELETE CASCADE,
  target_level integer NOT NULL,
  weight numeric,
  created_at timestamptz DEFAULT now(),
  UNIQUE (role_id, competency_id)
);

-- Ensure weight column exists (idempotent)
ALTER TABLE public.role_competencies
  ADD COLUMN IF NOT EXISTS weight numeric;

-- Enforce level range constraint if missing
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'role_competencies_target_level_chk'
  ) THEN
    ALTER TABLE public.role_competencies
      ADD CONSTRAINT role_competencies_target_level_chk
      CHECK (target_level BETWEEN 1 AND 4);
  END IF;
END$$;

CREATE INDEX IF NOT EXISTS idx_role_competencies_role ON public.role_competencies (role_id);
CREATE INDEX IF NOT EXISTS idx_role_competencies_comp ON public.role_competencies (competency_id);

-- =====================================================================================
-- USER PROFILES (basic profile + linkage to auth.users + selected current role)
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.user_profiles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid, -- optional linkage to auth.users
  full_name text,
  email text UNIQUE,
  current_role_id uuid REFERENCES public.roles(id),
  created_at timestamptz DEFAULT now()
);

-- Add FK to auth.users if not yet present
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'user_profiles_user_fk'
  ) THEN
    ALTER TABLE public.user_profiles
      ADD CONSTRAINT user_profiles_user_fk
      FOREIGN KEY (user_id) REFERENCES auth.users(id) ON DELETE SET NULL;
  END IF;
END$$;

CREATE INDEX IF NOT EXISTS idx_user_profiles_current_role ON public.user_profiles (current_role_id);

-- =====================================================================================
-- USER → COMPETENCIES with current assessed level + optional status tag
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.user_competencies (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_profile_id uuid NOT NULL REFERENCES public.user_profiles(id) ON DELETE CASCADE,
  competency_id uuid NOT NULL REFERENCES public.competencies(id) ON DELETE CASCADE,
  level integer,
  status_tag text, -- optional visual status for R/A/G traffic-light
  created_at timestamptz DEFAULT now(),
  UNIQUE (user_profile_id, competency_id)
);

-- Enforce level range, if present
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'user_competencies_level_chk'
  ) THEN
    ALTER TABLE public.user_competencies
      ADD CONSTRAINT user_competencies_level_chk
      CHECK (level IS NULL OR (level BETWEEN 1 AND 4));
  END IF;
END$$;

-- Enforce status tag enumerations
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'user_competencies_status_tag_chk'
  ) THEN
    ALTER TABLE public.user_competencies
      ADD CONSTRAINT user_competencies_status_tag_chk
      CHECK (status_tag IN ('red','amber','green') OR status_tag IS NULL);
  END IF;
END$$;

CREATE INDEX IF NOT EXISTS idx_user_competencies_user ON public.user_competencies (user_profile_id);
CREATE INDEX IF NOT EXISTS idx_user_competencies_comp ON public.user_competencies (competency_id);

-- =====================================================================================
-- Optional: USER → TARGET ROLES (allows tracking one or more targets)
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.user_target_roles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_profile_id uuid NOT NULL REFERENCES public.user_profiles(id) ON DELETE CASCADE,
  role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  priority integer, -- smaller=more preferred (optional)
  created_at timestamptz DEFAULT now(),
  UNIQUE (user_profile_id, role_id)
);

CREATE INDEX IF NOT EXISTS idx_user_target_roles_user ON public.user_target_roles (user_profile_id);
CREATE INDEX IF NOT EXISTS idx_user_target_roles_role ON public.user_target_roles (role_id);

-- =====================================================================================
-- ROLE ADJACENCY (similarity/transition score)
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.role_adjacency (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  source_role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  target_role_id uuid NOT NULL REFERENCES public.roles(id) ON DELETE CASCADE,
  score numeric,
  status_tag text, -- optional R/A/G per adjacency confidence or curation
  created_at timestamptz DEFAULT now(),
  UNIQUE (source_role_id, target_role_id)
);

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'role_adjacency_status_tag_chk'
  ) THEN
    ALTER TABLE public.role_adjacency
      ADD CONSTRAINT role_adjacency_status_tag_chk
      CHECK (status_tag IN ('red','amber','green') OR status_tag IS NULL);
  END IF;
END$$;

CREATE INDEX IF NOT EXISTS idx_role_adj_src ON public.role_adjacency (source_role_id);
CREATE INDEX IF NOT EXISTS idx_role_adj_tgt ON public.role_adjacency (target_role_id);
CREATE INDEX IF NOT EXISTS idx_role_adj_score ON public.role_adjacency (score);

-- =====================================================================================
-- LEARNING RESOURCES linked to competencies
-- =====================================================================================
CREATE TABLE IF NOT EXISTS public.learning_resources (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  competency_id uuid REFERENCES public.competencies(id) ON DELETE CASCADE,
  title text NOT NULL,
  url text NOT NULL,
  provider text,
  created_at timestamptz DEFAULT now(),
  UNIQUE (competency_id, url)
);

CREATE INDEX IF NOT EXISTS idx_learning_resources_comp ON public.learning_resources (competency_id);
CREATE INDEX IF NOT EXISTS idx_learning_resources_provider ON public.learning_resources (provider);

-- =====================================================================================
-- RLS Disablement for MVP seeding (re-enable later with ENABLE ROW LEVEL SECURITY)
-- =====================================================================================
ALTER TABLE public.users                 DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.roles                 DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.role_descriptions     DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.competencies          DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.role_competencies     DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_profiles         DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_competencies     DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_target_roles     DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.role_adjacency        DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.learning_resources    DISABLE ROW LEVEL SECURITY;

-- =====================================================================================
-- Helpful Views (optional): labels for numeric levels
-- =====================================================================================
-- Drop/create view safely
DO $$
BEGIN
  IF EXISTS (SELECT 1 FROM information_schema.views WHERE table_schema='public' AND table_name='level_labels') THEN
    DROP VIEW public.level_labels;
  END IF;
END$$;

CREATE VIEW public.level_labels AS
SELECT 1 AS level, 'Beginner'::text AS label
UNION ALL SELECT 2, 'Intermediate'
UNION ALL SELECT 3, 'Advanced'
UNION ALL SELECT 4, 'Master';
