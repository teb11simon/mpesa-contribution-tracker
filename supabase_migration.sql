-- ============================================================
-- M-Pesa Tracker Pro - Supabase Setup Migration
-- Run this in your Supabase SQL Editor (https://supabase.com/dashboard/project/nfegeltvychdouiywlxr/sql/new)
-- ============================================================

-- 1. Create the profiles table
CREATE TABLE IF NOT EXISTS public.profiles (
  id uuid references auth.users on delete cascade primary key,
  email text not null,
  role text default 'user'::text,
  status text default 'pending'::text,
  created_at timestamp with time zone default timezone('utc'::text, now()) not null
);

-- 2. Enable Row Level Security
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;

-- 3. Create policies
-- Allow users to read their own profile
CREATE POLICY "Users can read own profile"
  ON public.profiles
  FOR SELECT
  USING (auth.uid() = id);

-- Allow admin to read all profiles (via service_role key)
CREATE POLICY "Admin can read all profiles"
  ON public.profiles
  FOR SELECT
  USING (true);

-- Allow admin to update any profile (via service_role key)
CREATE POLICY "Admin can update profiles"
  ON public.profiles
  FOR UPDATE
  USING (true);

-- 4. Create the auto-profile trigger
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS trigger
LANGUAGE plpgsql
SECURITY DEFINER SET search_path = ''
AS $$
BEGIN
  INSERT INTO public.profiles (id, email, role, status)
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
    END
  );
  RETURN new;
END;
$$;

-- 5. Drop and recreate trigger
DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- ============================================================
-- Verify the setup
-- ============================================================
-- To check if it worked, run:
-- SELECT * FROM public.profiles;