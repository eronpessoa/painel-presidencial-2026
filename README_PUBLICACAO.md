# Painel Presidencial 2026 — publicação em nuvem V1.7

Este pacote contém a aplicação adaptada para Streamlit Community Cloud, mas **não contém** as bases eleitorais nem o GeoJSON, que devem ser copiados de seu computador.

## Estrutura e nomes obrigatórios

```
painel_presidencial_2026_cloud/
├── Scripts/
│   └── PAINEL_COMPARATIVO_PRESIDENCIAL_2026_CLOUD_V1_7.py
├── Resultados_1turno/
│   └── VOTOS_PRESIDENTE_MUNICIPIOS_1T.zip
├── Resultados_2turno/
│   └── RESULTADOS_MUNICIPAIS_2T_V1_1.sqlite3
├── Mapas_IBGE/
│   ├── municipios_br_leve.geojson
│   └── depara_tse_ibge_v15.csv
├── .streamlit/config.toml
├── requirements.txt
└── .gitignore
```

1. Copie o ZIP do 1º turno, o SQLite do 2º turno e o GeoJSON leve de 8,44 MB para as pastas acima, mantendo os nomes. O de-para já acompanha o pacote, mas suas correspondências propostas ainda precisam de validação oficial.
2. **Antes de publicar**, confirme que os arquivos contêm somente dados eleitorais destinados a divulgação pública. Não inclua senhas, dados pessoais não públicos, arquivos de trabalho internos ou o GeoJSON original de 676 MB.
3. No PowerShell, entre na pasta raiz deste pacote e teste:

   ```powershell
   python -m pip install -r requirements.txt
   python -m streamlit run Scripts/PAINEL_COMPARATIVO_PRESIDENCIAL_2026_CLOUD_V1_7.py
   ```

4. Crie um repositório GitHub (público somente depois de revisar os arquivos). Envie a **estrutura da pasta**, não a pasta contendo o pacote ZIP. Se os arquivos excederem os limites de upload do GitHub, use Git LFS ou outro mecanismo de armazenamento apropriado.
5. Entre em https://share.streamlit.io/ → Create app → selecione repositório e branch `main` → arquivo principal `Scripts/PAINEL_COMPARATIVO_PRESIDENCIAL_2026_CLOUD_V1_7.py` → Deploy. Selecione Python 3.12 nas configurações avançadas.
6. Configure a visibilidade pública nas configurações de compartilhamento e teste o endereço `.streamlit.app` em computador e celular.

## Observações

- O painel na nuvem é **somente leitura** e usa arquivos locais do repositório da aplicação; não sincroniza automaticamente com seu OneDrive nem com o TSE.
- O modo de demonstração e os campos de caminhos arbitrários foram desativados nesta edição pública.
- O painel local V1.7 original não é alterado.
- Os 13 vínculos do de-para são propostas de associação e requerem validação independente antes de divulgação definitiva.
- Dependências fixadas inicialmente para teste; caso haja diferenças visuais ou de compatibilidade, ajuste versões após testar em ambiente equivalente.
