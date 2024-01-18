# Automatizētā pieslēgšanās RTU sistēmām
![RTU_logo](https://github.com/patriksgrancarovs/ortus-automatizacija/assets/84973391/03074a15-90ac-4c5f-abd2-f0c74c816b8a)

## Projekta uzdevums

Šis projekts ir izstrādāts, lai atvieglotu lietotājam pieslēgties RTU sistēmām. Tas ietver ORTUS, E-studiju un nodarbību grafika atvēršanu.

## Izmantotās Python bibliotēkas un to izmantošanas iemesli

Projekts ir izstrādāts Python valodā un izmanto Selenium un Openpyxl bibliotēkas.
- **Selenium**: Šī bibliotēka tiek izmantota tīmekļa pārlūka automatizēšanai. Ar tās palīdzību ir iespējams atvērt pārlūku, ievadīt datus un veikt dažādas darbības, kā piemēram noklikšķināt uz pogām vai ievadīt tekstu.
- **Openpyxl**: Šī bibliotēka tiek izmantota Excel datņu apstrādei Python valodā. Tā atļauj apstrādāt Excel failu, kuru šajā projektā izmanto, lai saglabātu lietotāja datus.

Izmantotās pamatbibliotēkas
- **os**: Šo bibliotēku izmantojām, lai Python var darboties ar failu sistēmu, piemēram, lasīt, rakstīt un dzēst failus, kā arī notīrīt termināli.
- **sys**: Bibliotēka ļauj Python programmai beigt darbību.
- **time**: Šī bibliotēka tiek izmantota, lai darbotos ar laika funkcijām. Tas ļauj Python programmai aizkavēt izpildi.

## Programmatūras izmantošana

Lai izmantotu šo programmatūru, jāielādē AutoRTU.exe programma un jāpalaiž, nav svarīga tās atrašanās vieta datorā, taču atrašanās vietā tiks izveidota Excel datne ar lietotāja datiem. Tālāk būs nepieciešams ievadīt savu ORTUS lietotājvārdu un paroli, kā arī jānorāda jūsu programmas kods, kurss un grupa. Pēc tam variet izvēlēties, kurai sistēmai vēlaties pieslēgties. Programma automātiski atvērs pārlūku un pieslēgsies izvēlētajai sistēmai. Var pieslēgties vairākām sistēmām secīgi. Pēc darba beigšanas var nospiest pogu **4** lai izdzēstu lietotāja datus, vai **5** lai beigtu darbu saglabājot lietotāja datus. Piezīme: programma paroli nesaglabā drošības apsvērumu dēļ, tā katru reizi no jauna jāieraksta restartējot programmu.

[Video ar programmas darbību](https://www.youtube.com/watch?v=udea_gXEAQE)

## Autori
- Patriks Grančarovs
- Jurģis Pāvulītis

