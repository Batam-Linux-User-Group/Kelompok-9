function hitungRataRata(nilai) {
    let total = 0;

    for (let i = 0; i < nilai.length; i++) {
        total = total + nilai[i];
    }

    return total / nilai.length;
}

let nilaiMahasiswa = [80, 75, 90, 85, 70];

let rataRata = hitungRataRata(nilaiMahasiswa);

console.log("Rata-rata nilai:", rataRata);