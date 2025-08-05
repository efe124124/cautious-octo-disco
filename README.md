# cautious-octo-disco

Bu depo, bir Excel dosyasındaki tüm verileri satır ve sütun düzenini bozmadan
başka bir Excel dosyasına aktarmak için kullanılabilecek basit bir Python
betiği içerir.

## Kullanım

1. Gerekli bağımlılığı kurun:

   ```bash
   pip install openpyxl
   ```

2. Betiği çalıştırın:

   ```bash
   python excel_copy.py kaynak.xlsx hedef.xlsx
   ```

   `kaynak.xlsx` dosyasındaki tüm sayfalar satır ve sütunları değiştirilmeden
   `hedef.xlsx` dosyasına kopyalanır.

