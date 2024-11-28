# Zasady programowania obiektowego

## Kluczowe pojęcia w programowaniu obiektowym

Programowanie obiektowe (OOP) to podejście do programowania, które pozwala organizować kod w sposób zorientowany na obiekty. Oto kluczowe pojęcia:

1. **Klasa**: Szablon, który definiuje strukturę danych (właściwości) i funkcje (metody), które będą dostępne w obiektach utworzonych na jej podstawie.
2. **Obiekt**: Instancja klasy, która przechowuje konkretne dane i wykonuje operacje zdefiniowane w klasie.
3. **Właściwości**: Zmienne, które przechowują dane związane z obiektem.
4. **Metody**: Funkcje, które operują na danych obiektu i wykonują zadania.
5. **Dziedziczenie**: Mechanizm, który pozwala jednej klasie przejmować właściwości i metody innej klasy.
6. **Polimorfizm**: Możliwość wywoływania tej samej metody w różnych klasach, ale z różnym zachowaniem.

---

## Tworzenie klas i obiektów

### Klasy

Klasa definiuje strukturę danych i zachowania, które będą wspólne dla wszystkich obiektów utworzonych na jej podstawie.

Przykład ogólny:

- Klasa `Samochod` może zawierać właściwości, takie jak `marka`, `model`, `kolor` oraz metody, takie jak `start` i `zatrzymaj`.

### Obiekty

Obiekt jest instancją klasy i przechowuje konkretne dane. Na przykład:

- Obiekt `mojSamochod` może mieć `marka = "Toyota"`, `model = "Corolla"`, `kolor = "czerwony"`.

---

## Właściwości i metody

### Właściwości

Właściwości to zmienne przechowywane w obiekcie. Mogą być używane do przechowywania danych, takich jak parametry lub ustawienia.

Przykład:

- `Imie` w klasie `Osoba`.

### Metody

Metody to funkcje, które wykonują operacje na danych obiektu. Na przykład:

- `PrzedstawSie()` w klasie `Osoba`, która wyświetla dane o osobie.

---

## Dziedziczenie

Dziedziczenie pozwala jednej klasie (klasie podrzędnej) przejmować właściwości i metody innej klasy (klasy nadrzędnej). Dzięki temu możemy unikać duplikacji kodu.

Przykład:

- Klasa `Pojazd` ma właściwości `typ` i `kolor`. Klasa `Samochod` może dziedziczyć te właściwości i dodać swoje własne, takie jak `liczbaDrzwi`.

---

## Polimorfizm

Polimorfizm pozwala używać tej samej metody w różnych klasach, ale z różnym zachowaniem. Dzięki temu można tworzyć elastyczne i uniwersalne rozwiązania.

Przykład:

- Metoda `Zatrzymaj` w klasie `Pojazd` może działać inaczej w klasie `Samochod` i w klasie `Rower`.

---

## Zalety programowania obiektowego

1. **Modularność**: Kod jest podzielony na mniejsze, niezależne części (klasy i obiekty).
2. **Reużywalność**: Klasy i obiekty można łatwo wykorzystać w innych projektach.
3. **Łatwiejsza konserwacja**: Kod jest bardziej przejrzysty i łatwy do zarządzania.
4. **Elastyczność**: Możliwość dostosowywania kodu dzięki dziedziczeniu i polimorfizmowi.

---

## Podsumowanie

Programowanie obiektowe to jeden z fundamentów nowoczesnego programowania. Dzięki zrozumieniu OOP możesz tworzyć lepiej zorganizowany, bardziej elastyczny i wydajniejszy kod, który łatwiej będzie rozwijać i utrzymywać.
