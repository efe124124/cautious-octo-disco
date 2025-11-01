# Excel Portföy Takip Aracı

Bu depo, canlı piyasa verilerini kullanarak Excel üzerinde portföy takibi yapmanızı
sağlayan örnek bir otomasyon içerir. `portfolio_tracker.py` betiği, bir JSON
konfigürasyon dosyasını okur, Yahoo Finance üzerinden fiyat/kurları çeker ve
portföyünüzü özetleyen biçimlendirilmiş bir Excel çalışma kitabı üretir.

## Özellikler

- BIST ve uluslararası semboller için son fiyat, günlük değişim ve hacim bilgisi
  (Yahoo Finance üzerinden) toplanır.
- Farklı para birimlerinde tutulan varlıklar, seçtiğiniz temel para birimine FX
  kuruyla çevrilir.
- Toplam değer, maliyet, gerçekleşmemiş kâr/zarar, günlük değişim ve portföy ağırlıkları
  otomatik hesaplanır.
- Excel dosyası içinde "Holdings" (varlıklar) ve "Summary" (özet) olarak iki sayfa
  oluşturulur; tablo filtreleri, dondurulmuş başlıklar ve yüzde/para biçimleri hazır gelir.

## Gereksinimler

Python 3.10+ sürümü önerilir. Gerekli kütüphaneleri kurmak için:

```bash
python -m venv .venv
source .venv/bin/activate  # Windows için: .venv\Scripts\activate
pip install -r requirements.txt
```

## Konfigürasyon

`portfolio_config.json` dosyasında temel para biriminizi (`base_currency`) ve
varlık listenizi tanımlayabilirsiniz. Örnek yapı:

```json
{
  "base_currency": "TRY",
  "holdings": [
    { "symbol": "SASA.IS", "quantity": 120, "cost_basis": 33.45, "currency": "TRY" },
    { "symbol": "THYAO.IS", "quantity": 85, "cost_basis": 78.10, "currency": "TRY" },
    { "symbol": "TSLA", "quantity": 4, "cost_basis": 180.0, "currency": "USD" },
    { "symbol": "GLD", "quantity": 10, "cost_basis": 175.25, "currency": "USD" }
  ]
}
```

- `symbol`: Yahoo Finance sembolü. BIST hisseleri için `.IS` uzantısını unutmayın.
- `quantity`: Lot/adet sayısı (ondalıklı miktar da olabilir).
- `cost_basis`: Birim alış maliyeti (varlığın kendi para birimi cinsinden).
- `currency`: Varlığın işlem para birimi. Belirtmezseniz `base_currency` kabul edilir.

## Kullanım

```bash
python portfolio_tracker.py --config portfolio_config.json --output portfoy.xlsx
```

- `--config`: Portföy tanımını içeren JSON dosyası.
- `--output`: Oluşturulacak Excel dosyasının adı (varsayılan `portfolio.xlsx`).

Komut çalıştıktan sonra belirtilen Excel dosyası oluşur. Excel’de **Veri > Tümünü
Yenile** komutuyla dosyayı tekrar oluşturmanıza gerek kalmadan en son haliyle
görüntüleyebilirsiniz. Portföyünüz güncellendikçe JSON dosyasını düzenleyip aynı
komutu yeniden çalıştırmanız yeterli.

## Sık Karşılaşılan Sorular

- **Veriler ne kadar gerçek zamanlı?** Yahoo Finance genellikle birkaç dakikalık
  gecikmeli fiyatlar sunar. Resmî veya ücretli API’ler gerektiğinde Power Query
  ya da farklı kütüphaneler entegre edilebilir.
- **Döviz kurları gelmezse ne yapmalıyım?** Seçilen para birimi çifti Yahoo Finance
  üzerinde bulunamayabilir. Örneğin TRY tabanlı olmayan bir para birimini desteklemek
  için `USDTRY=X`, `EURTRY=X` gibi sembollerin mevcut olduğundan emin olun.
- **Daha sık otomatik güncelleme istiyorum.** Betiği Windows Görev Zamanlayıcı,
  cron veya Power Automate üzerinden belirli aralıklarla çalıştırabilirsiniz.

## Lisans

Bu depo örnek amaçlıdır, herhangi bir lisans belirtilmemiştir.
