# Oftalmocastro – Controle de Exames e Cirurgias

Versão React + Vite com:

- Controle mensal de exames e cirurgias
- Campo de WhatsApp ao lado do paciente, com botão direto para abrir a conversa
- Campo de observações ampliado e observações visíveis nas tabelas
- Dashboard mensal com análise automática
- Exportação real em Excel `.xlsx` com abas **Exames**, **Cirurgias** e **Resumo Mensal**
- Filtros por mês, médico, colaborador e status

## Executar

```bash
npm install
npm run dev
```

> Os dados ainda são armazenados no `localStorage`. Para uso compartilhado por vários computadores, o próximo passo é conectar o React a um banco de dados/API.


## Novidades V3
- Cadastros dinâmicos de exames, profissionais e procedimentos.
- Botão "Cadastros da clínica" para incluir novos itens manualmente.
- Resultados reorganizados em blocos e com análise mensal.
- Os novos itens passam a aparecer automaticamente nos formulários e filtros.
