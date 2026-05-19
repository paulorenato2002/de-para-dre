-- Estrutura base para o mini banco de parametrização.
-- O app usa a service key do Supabase, então pode ler e gravar mesmo com RLS ligado.

create table if not exists public.empresas (
    id text primary key,
    nome text not null,
    ativo boolean not null default true
);

create table if not exists public.parametrizacao_contas (
    id bigserial primary key,
    empresa_id text not null references public.empresas(id) on delete cascade,
    conta_contabil text not null,
    fornecedor_cliente text not null,
    created_at timestamptz not null default now()
);

create index if not exists idx_parametrizacao_contas_empresa_conta
    on public.parametrizacao_contas (empresa_id, conta_contabil);

insert into public.empresas (id, nome, ativo) values
    ('401', '401 - SST', true),
    ('370', '370 - STI', true),
    ('536', '536 - SCD', true),
    ('570', '570 - SSD', true),
    ('556', '556 - SAC', true)
on conflict (id) do update
set nome = excluded.nome,
    ativo = excluded.ativo;

-- Se quiser começar já com a empresa 401, pode inserir a parametrização pela tela do app
-- ou rodar um seed com os pares Conta Débito x Cód. Plano de conta nº. 04.
