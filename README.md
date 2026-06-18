# NexHotel

Aplicativo desktop (Electron) para gestao completa de hoteis, pousadas, hostels e operacoes de hospedagem.

## Funcionalidades

- Dashboard completo com KPIs e alertas operacionais
- Gestao de quartos (disponivel, ocupado, manutencao)
- Controle de reservas com fluxo agendada -> confirmada -> hospedado -> check-out
- Cadastro de clientes/hospedes
- Financeiro completo (receitas, despesas, metodos de pagamento, fechamento de caixa)
- Controle de estoque com entradas, saidas e alerta de minimo
- Cadastro de funcionarios
- Registro de ponto
- Ordens de servico (limpeza, manutencao, atendimento, alimentacao)
- Controle de contas bancarias
- Relatorios com filtros de periodo
- Notificacoes internas
- Auditoria e historico de acoes
- App/Painel do hospede dentro do mesmo sistema
- App/Painel de funcionarios dentro do mesmo sistema

## Armazenamento local

Os dados ficam em:

- `%APPDATA%/NexHotel/data/nexhotel_db.json`

## Como rodar

1. Instale Node.js 18+
2. No terminal dentro da pasta do projeto:

```bash
npm install
npm run start
```

## Build instalador Windows

```bash
npm run build
```

Saida:

- Pasta `dist_installer`
- Instalador `.exe` (NSIS)

## Observacao sobre a logo

A logo do NexHotel esta embutida no projeto:

- `assets/nexhotel-logo.png`
