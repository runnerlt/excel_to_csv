# Excel failo konverteris į CSV
## Informacija  vartotojui
Programa apjungia visus `input` kataloge esančius `.xlsx` failus į vieną `.csv` failą. Rezultatas išsaugomas `output` kataloge. Sugeneruojamas `csv` failas, kurio pavadinime yra data ir laikas, kada buvo paleista programa.

Kiekvienas programos paleidimas sugeneruoja naują `csv` failą, kurio turinyje turime tuo metu `input` kataloge buvusių `excel` failų agreguotą bei dedublikuotą turinį.

## Dedublikavimas
Duomenys dedublikuojami pagal pirmojo stulpelio reikšmes `.xlsx` failuose. Likę stulpeliai imami tie, kurie turi ilgiausią tekstą pagal tą dedublikavimo stulpelio vertę.

**Pavizdys**
| Kodas       | Pavadinimas                  |
|-------------|------------------------------|
| Kirkutytė   | test deduplicate             |
| Kirkutytė   | test deduplicate functionality|

Turėtume vieną įrašą:
| Kodas       | Pavadinimas                   |
|-------------|-------------------------------|
| Kirkutytė   | test deduplicate functionality|



## Skyriklis
Kaip skyriklį csv faile naudojame `;`.
Jei norite pakeisti skyriklį, tai galite padaryti and failo `convert.cmd` paspaudę `show more options`, tada `edit`, ir pakeitę `";"` vertę į savo norimą.

## Programos naudojimas
- Sudėkite konvertuotinus excel failus į `input` katalogą.
- Paleiskite `convert.cmd` failą.


# Informacija programuotojui
Siekiamybė yra pateikti vartotojui tokią programą, kuri galėtų veikti standartinėje `Windows` instaliacijoje be jokių papildomų įdiegtų programų. Daroma prielaida, kad vartotojas nėra IT srities atstovas, taigi, jam nėra žinoma kas yra `Python` ar kažkokie papildomi instaliaciniai paketai. Taipogi mes dažnai turime situaciją, kai vartotojas turi apribotas teises savo kompiuteryje, todėl negali įdiegti papildomų programų. Tas lemia pakankamai aiškų apribojimą, kad programa turi veikti `stand-alone` režimu.

Kad tai pasiekti, mes supakuojame Python programą naudodami `PyInstaller` biblioteką. Tai leidžia mums sukurti vieną `.exe` failą, kuris veiks be jokių papildomų instaliacijų.

Kad sugeneruoti `.exe` failą, naudojame:
```py
pyinstaller --onefile excel_to_csv_batch.py
```

Naudojimas:
```cmd
excel_to_csv_batch.exe "input" ";"
```