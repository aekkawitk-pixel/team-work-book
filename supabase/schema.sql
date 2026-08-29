-- ============================================================
--  สมุดคุมงานทีม (Team Work Book) — โครงสร้างฐานข้อมูล
--  รันไฟล์นี้ทั้งไฟล์ใน Supabase SQL Editor บนโปรเจกต์เปล่า
-- ============================================================

create extension if not exists pgcrypto;

-- ---------- settings (1 แถวต่อผู้ใช้) ----------
create table public.settings (
  owner_id uuid primary key references auth.users(id) on delete cascade,
  display_name text not null default '',
  theme text not null default 'system' check (theme in ('light','dark','system')),
  default_due_days text not null default '',
  default_report_cycle text not null default '',
  recent jsonb not null default '{"member":[],"project":[],"requester":[]}'::jsonb,
  updated_at timestamptz not null default now()
);

-- ---------- members (รวมตัวหัวหน้าเองด้วย is_self) ----------
create table public.members (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  name text not null,
  nickname text not null default '',
  active boolean not null default true,
  is_self boolean not null default false,
  sort_order int not null default 0,
  created_at timestamptz not null default now()
);
create unique index members_one_self on public.members(owner_id) where is_self;

-- ---------- requesters (ผู้มอบหมายงาน) ----------
create table public.requesters (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  name text not null,
  active boolean not null default true,
  sort_order int not null default 0,
  created_at timestamptz not null default now()
);

-- ---------- projects ----------
create table public.projects (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  name text not null,
  description text not null default '',
  status text not null default 'active' check (status in ('active','done')),
  start_at date,
  end_at date,
  archived boolean not null default false,
  sort_order int not null default 0,
  created_at timestamptz not null default now()
);

-- ---------- tasks ----------
create table public.tasks (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  title text not null,
  description text not null default '',
  requester_id uuid references public.requesters(id) on delete set null,
  assignee_id uuid references public.members(id) on delete set null,
  project_id uuid references public.projects(id) on delete set null,
  status text not null default 'received'
    check (status in ('received','in_progress','blocked','completed','paused')),
  priority text not null default 'normal' check (priority in ('high','normal','low')),
  received_at timestamptz not null default now(),
  due_at date,
  report_cycle text not null default ''
    check (report_cycle in ('','weekly','biweekly','monthly')),
  repeat_next boolean not null default false,
  completed_at date,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index tasks_owner_status   on public.tasks(owner_id, status);
create index tasks_owner_due      on public.tasks(owner_id, due_at);
create index tasks_owner_project  on public.tasks(owner_id, project_id);
create index tasks_owner_assignee on public.tasks(owner_id, assignee_id);

-- ---------- task_collaborators (ผู้ร่วมงาน) ----------
create table public.task_collaborators (
  task_id uuid not null references public.tasks(id) on delete cascade,
  member_id uuid not null references public.members(id) on delete cascade,
  owner_id uuid not null references auth.users(id) on delete cascade,
  primary key (task_id, member_id)
);
create index task_collab_member on public.task_collaborators(owner_id, member_id);

-- ---------- task_updates (ประวัติความคืบหน้า) ----------
create table public.task_updates (
  id uuid primary key default gen_random_uuid(),
  task_id uuid not null references public.tasks(id) on delete cascade,
  owner_id uuid not null references auth.users(id) on delete cascade,
  message text not null default '',
  status text not null default ''
    check (status in ('','received','in_progress','blocked','completed','paused')),
  created_at timestamptz not null default now()
);
create index task_updates_task on public.task_updates(task_id, created_at);

-- ---------- notes (บันทึกของหัวหน้า) ----------
create table public.notes (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  note_date date not null default current_date,
  body text not null default '',
  project_id uuid references public.projects(id) on delete set null,
  created_at timestamptz not null default now()
);
create index notes_owner_date on public.notes(owner_id, note_date);

-- ---------- week_notes (แผน / ปัญหา รายสัปดาห์) ----------
create table public.week_notes (
  owner_id uuid not null references auth.users(id) on delete cascade,
  week_start date not null,
  plan text not null default '',
  risk text not null default '',
  updated_at timestamptz not null default now(),
  primary key (owner_id, week_start)
);

-- ---------- trigger: updated_at ----------
create or replace function public.touch_updated_at()
returns trigger language plpgsql as $$
begin new.updated_at = now(); return new; end $$;

create trigger tasks_touch      before update on public.tasks
  for each row execute function public.touch_updated_at();
create trigger settings_touch   before update on public.settings
  for each row execute function public.touch_updated_at();
create trigger week_notes_touch before update on public.week_notes
  for each row execute function public.touch_updated_at();

-- ---------- RLS : ทุกตารางเห็นเฉพาะข้อมูลของตัวเอง ----------
alter table public.settings           enable row level security;
alter table public.members            enable row level security;
alter table public.requesters         enable row level security;
alter table public.projects           enable row level security;
alter table public.tasks              enable row level security;
alter table public.task_collaborators enable row level security;
alter table public.task_updates       enable row level security;
alter table public.notes              enable row level security;
alter table public.week_notes         enable row level security;

do $$
declare t text;
begin
  foreach t in array array['settings','members','requesters','projects','tasks',
                           'task_collaborators','task_updates','notes','week_notes']
  loop
    execute format($f$
      create policy "own_select" on public.%I for select to authenticated
        using (owner_id = (select auth.uid()));
      create policy "own_insert" on public.%I for insert to authenticated
        with check (owner_id = (select auth.uid()));
      create policy "own_update" on public.%I for update to authenticated
        using (owner_id = (select auth.uid()))
        with check (owner_id = (select auth.uid()));
      create policy "own_delete" on public.%I for delete to authenticated
        using (owner_id = (select auth.uid()));
    $f$, t, t, t, t);
  end loop;
end $$;

-- ---------- ตั้งค่าเริ่มต้นอัตโนมัติเมื่อมีผู้ใช้ใหม่ ----------
create or replace function public.init_workbook()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  insert into public.settings(owner_id, display_name)
    values (new.id, coalesce(new.raw_user_meta_data->>'display_name',''))
    on conflict (owner_id) do nothing;
  insert into public.members(owner_id, name, is_self, sort_order)
    values (new.id,
            coalesce(nullif(new.raw_user_meta_data->>'display_name',''),'หัวหน้าแผนก'),
            true, -1);
  return new;
end $$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.init_workbook();
