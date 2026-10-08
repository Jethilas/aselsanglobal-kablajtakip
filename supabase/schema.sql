-- Kablaj Hata Oran Takip — Supabase şeması
-- Supabase Dashboard → SQL Editor → New query → yapıştır → Run.
-- Roller: izleyici (varsayılan) / yonetici / admin

-- 1) Profiller tablosu
create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text,
  rol text not null default 'izleyici'
    check (rol in ('admin', 'yonetici', 'izleyici')),
  created_at timestamptz default now()
);

-- 2) Yeni kayıt olana otomatik profil (rol = izleyici)
create or replace function public.profil_olustur()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email, rol)
  values (new.id, new.email, 'izleyici')
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists trg_profil_olustur on auth.users;
create trigger trg_profil_olustur
  after insert on auth.users
  for each row execute function public.profil_olustur();

-- 3) Satır güvenliği (RLS)
alter table public.profiles enable row level security;

drop policy if exists "kendi profilini okur" on public.profiles;
create policy "kendi profilini okur"
  on public.profiles for select
  using (auth.uid() = id);

drop policy if exists "admin hepsini okur" on public.profiles;
create policy "admin hepsini okur"
  on public.profiles for select
  using (exists (
    select 1 from public.profiles p
    where p.id = auth.uid() and p.rol = 'admin'
  ));

drop policy if exists "admin rol degistirir" on public.profiles;
create policy "admin rol degistirir"
  on public.profiles for update
  using (exists (
    select 1 from public.profiles p
    where p.id = auth.uid() and p.rol = 'admin'
  ));

-- 4) İlk yöneticiyi yükselt (e-postayı değiştirip çalıştır):
-- update public.profiles set rol = 'admin' where email = 'sen@ornek.com';
-- Rol düşür/yükselt örnekleri:
-- update public.profiles set rol = 'yonetici' where email = 'sorumlu@ornek.com';
