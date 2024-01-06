import math
import os
import sys
import pandas as pd
from tkinter import filedialog, Tk, Button, Label, Entry, messagebox


# Dosya seçme fonksiyonu
def dosya_sec():
    dosya_yolu = filedialog.askopenfilename(title="Lütfen bir Excel dosyası seçin", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not dosya_yolu:
        messagebox.showwarning("Uyarı", "Dosya seçmediniz! Program sonlandırılıyor.")
        print("Dosya seçilmedi. Program sonlandırılıyor.")
        sys.exit()
    return dosya_yolu

# K satısını kullanıcıdan alma fonksiyonu
def k_sayisi_penceresi():
    pencere = Tk()
    pencere.title("K Sayısı Belirleme")
    pencere.geometry("300x150")

    label = Label(pencere, text="Lütfen k sayısını giriniz:")
    label.pack(pady=10)

    k_giris = Entry(pencere)
    k_giris.pack(pady=10)

    def k_sayisi_al():
        try:
            global k
            k = int(k_giris.get())
            pencere.destroy()
        except ValueError:
            messagebox.showerror("Hata", "Geçersiz giriş. Lütfen bir sayı girin.")

    button = Button(pencere, text="Onayla", command=k_sayisi_al)
    button.pack(pady=10)

    pencere.protocol("WM_DELETE_WINDOW", sys.exit) # pencereyi kapattığımda hata almamak için kodu sonlandırıyorum

    pencere.mainloop()

# K_means kümeleme fonksiyonum, fonksiyonum genel hatlarıyla 4 aşamadan oluşuyor
def k_means(dataset):
    while (True):

        #1- her merkezin datasetimdeki verilere uzaklıkları bir sözlüğe alınıyor, en son sırasıyla tek bir sözlükte toplanıyor
        uzaklik_dict = {}
        for i in range(k):
            uzaklik_dict_temp = {}
            for j in range(
                    len(dataset)):
                birinci_sutun_fark = dataset.iloc[j, 0] - merkezler.iloc[i, 0] # iloc[satır, sütun]
                ikinci_sutun_fark = dataset.iloc[j, 1] - merkezler.iloc[i, 1]
                uzaklik = math.sqrt(pow(birinci_sutun_fark, 2) + pow(ikinci_sutun_fark, 2)) # uzaklıklar öklid uzaklık formülüyle hesaplanır
                uzaklik_dict_temp.update({j: uzaklik})
            uzaklik_dict.update({i: uzaklik_dict_temp})

        #2- uzaklık sözlüğündeki mesafeler karşılaştırılır ve en kısa mesafesi olan merkez, ilgili noktanın kümesi olarak seçilir
        kume_no = []
        for j in range(len(dataset)):
            min_mesafe = []
            for i in range(k):
                min_mesafe.append(uzaklik_dict[i][j])
            mim_deger = min(min_mesafe)
            kume_no.append(min_mesafe.index(mim_deger))
        dataset['Küme'] = kume_no   # belirlenen küme numaraları, dataframe de oluşturulan yeni sütuna sırasına göre eklenir


        #3- yeni merkezler hesaplanır
        yeni_birinci_kolon = []
        yeni_ikinci_kolon = []
        for i in range(k):
            birinci_kolon = []
            ikinci_kolon = []
            for j in range(len(dataset)):
                if dataset.iloc[j, 2] == i:
                    birinci_kolon.append(dataset.iloc[j, 0]) # iloc[satır, sütun]
                    ikinci_kolon.append(dataset.iloc[j, 1])

            if len(birinci_kolon) != 0:
                yeni_birinci_kolon.append(sum(birinci_kolon) / len(birinci_kolon)) #\
            else:                                                                   #\
                yeni_birinci_kolon.append(0)                                         #> sıfıra bölünme hatasını engelledim
            if len(ikinci_kolon) != 0:                                             #/
                yeni_ikinci_kolon.append(sum(ikinci_kolon) / len(ikinci_kolon))    #/
            else:
                yeni_ikinci_kolon.append(0)


        #4- yeni merkezler eski merkezlerle aynı mı diye kontrol edilir, aynıysa döngüden çıkılır ve işlem biter
        eski_birinci_kolon = merkezler.iloc[:, 0].tolist()
        eski_ikinci_kolon = merkezler.iloc[:, 1].tolist()

        if yeni_birinci_kolon == eski_birinci_kolon and yeni_ikinci_kolon == eski_ikinci_kolon:
            break
        else:
            merkezler.iloc[:, 0] = yeni_birinci_kolon
            merkezler.iloc[:, 1] = yeni_ikinci_kolon
            continue


# kodun çalıştırılması -------------------------------------

dosya_yolu = dosya_sec()                                # seçilen dosya yolu kaydedildi
dataset = pd.read_excel(dosya_yolu)                     # ilgili excel dosyası dataframe olarak kaydedildi
k_sayisi_penceresi()                                    # kullanıcıdan k sayısı alındı
merkezler = dataset.sample(k).reset_index(drop=True)    # başlangıç için rastgele k adet merkez belirlendi
k_means(dataset)                                        # k_means kümeleme fonksiyonu çalıştırıldı

output_file_path = os.path.join(os.path.dirname(dosya_yolu), 'dataset_k_means_Çıktı.xlsx')  # çıktının konumu ve adı belirlendi
dataset.to_excel(output_file_path, index=False)  # çıktı excel dosyasına dönüştürülüp ilgili konuma kaydedildi

messagebox.showinfo("İşlem Başarılıyla Gerçekleştirildi", f"Sonuç dosyası '{output_file_path}' olarak kaydedildi.")# işlem başarılı mesajı




