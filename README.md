# AKT – Excel → Word (Streamlit)

Bu app **Excel (.xlsx)** faylından `Satış sıralaması` və `siyahı` sütunlarını oxuyur,
seçilmiş satış nömrələri üzrə **tək bir Word (.docx)** sənədində aşağıdakı kimi sətirlər yaradır:
```
1-ci NV: 125, 150, 434
2-ci NV: 201, 202
...
```
və şablondakı placeholder mətni olan paraqrafa/cədvələ yerləşdirir.

## İstifadə
1. **Excel (.xlsx)** faylını yüklə (vərəq adını boş buraxsan 1-ci vərəq oxunur).
2. **Word (.docx) şablon** yüklə — içində mətn olaraq bu yazılardan biri olmalıdır:
   - NETICELER VE SIYAHI BURA YAZILACAQ (diakritika variantları da dəstəklənir)
3. NV satış nömrələrini vergüllə daxil et: `1,2,3`.
4. **AKT yarad və endir** düyməsinə bas.

## Lokal
```
pip install -r requirements.txt
streamlit run app.py
```
