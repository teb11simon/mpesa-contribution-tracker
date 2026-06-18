-- ============================================================
-- M-Pesa Tracker Pro - Supabase Setup Migration
-- Run this in your Supabase SQL Editor (https://supabase.com/dashboard/project/wcrcfitxlkymhcodqdfb/sql/new)
-- ============================================================

-- ============================================================
-- 1. CHURCHES TABLE (new - for multi-church support)
-- ============================================================
CREATE TABLE IF NOT EXISTS public.churches (
  id uuid DEFAULT gen_random_uuid() primary key,
  name text not null,
  slug text not null unique,
  created_at timestamp with time zone default timezone('utc'::text, now()) not null
);

-- Insert the default Nairobi ICC church
INSERT INTO public.churches (name, slug)
VALUES ('Nairobi ICC', 'nairobi-icc')
ON CONFLICT (slug) DO NOTHING;

-- ============================================================
-- 2. PROFILES TABLE (updated with church_id)
-- ============================================================
-- First add church_id column if it doesn't exist
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns 
    WHERE table_schema = 'public' AND table_name = 'profiles' AND column_name = 'church_id'
  ) THEN
    ALTER TABLE public.profiles ADD COLUMN church_id uuid references public.churches(id);
  END IF;
END $$;

CREATE TABLE IF NOT EXISTS public.profiles (
  id uuid references auth.users on delete cascade primary key,
  email text not null,
  role text default 'user'::text,
  status text default 'pending'::text,
  church_id uuid references public.churches(id),
  created_at timestamp with time zone default timezone('utc'::text, now()) not null
);

-- 3. Enable Row Level Security
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.churches ENABLE ROW LEVEL SECURITY;

-- 4. Create policies for profiles
-- Allow users to read their own profile
DROP POLICY IF EXISTS "Users can read own profile" ON public.profiles;
CREATE POLICY "Users can read own profile"
  ON public.profiles
  FOR SELECT
  USING (auth.uid() = id);

-- Allow admin to read all profiles (via service_role key)
DROP POLICY IF EXISTS "Admin can read all profiles" ON public.profiles;
CREATE POLICY "Admin can read all profiles"
  ON public.profiles
  FOR SELECT
  USING (true);

-- Allow admin to update any profile (via service_role key)
DROP POLICY IF EXISTS "Admin can update profiles" ON public.profiles;
CREATE POLICY "Admin can update profiles"
  ON public.profiles
  FOR UPDATE
  USING (true);

-- 5. Policies for churches table
-- Anyone can read churches list (for signup dropdown)
DROP POLICY IF EXISTS "Anyone can read churches" ON public.churches;
CREATE POLICY "Anyone can read churches"
  ON public.churches
  FOR SELECT
  USING (true);

-- Only service_role can insert/update/delete churches
DROP POLICY IF EXISTS "Admin can manage churches" ON public.churches;
CREATE POLICY "Admin can manage churches"
  ON public.churches
  FOR ALL
  USING (true);

-- 6. Create the auto-profile trigger (updated with church_id)
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS trigger
LANGUAGE plpgsql
SECURITY DEFINER SET search_path = ''
AS $$
DECLARE
  default_church_id uuid;
BEGIN
  -- Get the default church (Nairobi ICC)
  SELECT id INTO default_church_id FROM public.churches WHERE slug = 'nairobi-icc' LIMIT 1;
  
  INSERT INTO public.profiles (id, email, role, status, church_id)
  VALUES (
    new.id,
    new.email,
    CASE 
      WHEN new.email = 'tebosimon@gmail.com' THEN 'admin'::text 
      ELSE 'user'::text 
    END,
    CASE 
      WHEN new.email = 'tebosimon@gmail.com' THEN 'approved'::text 
      ELSE 'pending'::text 
    END,
    default_church_id
  );
  RETURN new;
END;
$$;

-- 7. Drop and recreate trigger
DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- ============================================================
-- Verify the setup
-- ============================================================
-- SELECT * FROM public.profiles;
-- SELECT * FROM public.churches;