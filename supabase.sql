-- Ejecuta esto en Supabase: Dashboard → SQL Editor → New query → pega y "Run".

create table if not exists wedding_data (
  id text primary key,
  data jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

-- La app usa un único registro (id = 'main') que guarda todo el estado como JSON.
insert into wedding_data (id, data)
values ('main', '{}'::jsonb)
on conflict (id) do nothing;

-- RLS: como esta app no tiene login propio, dejamos lectura/escritura abiertas
-- para cualquiera que tenga la URL y la anon key (ambas quedan visibles en el
-- código del sitio, como es normal en Supabase). Es aceptable para un uso
-- privado de pareja, pero no publiques el link ni el repo con datos reales.
alter table wedding_data enable row level security;

create policy "wedding_data_select" on wedding_data for select using (true);
create policy "wedding_data_insert" on wedding_data for insert with check (true);
create policy "wedding_data_update" on wedding_data for update using (true);

-- Habilita Realtime para que los cambios se vean en vivo en ambos dispositivos.
alter publication supabase_realtime add table wedding_data;
