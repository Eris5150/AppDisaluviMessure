🪚 CortesAluminioWpf

A C# WPF (.NET) desktop app for planning and optimizing aluminum cuts.
It helps minimize waste by suggesting the best cutting strategy for each stock bar.

🚀 Key Features

Define stock bar length (e.g., 6m).

Add and manage required pieces (sorted automatically).

Visual status of pieces: 🟩 In Play | 🟥 Used | ⬜ Excluded.

Save completed cutting lines and export results to Excel.

🤖 Cutting Algorithm

The core algorithm compares:

Best single piece ≤ remaining length.

Best pair of pieces ≤ remaining length.

It selects the option that produces the smallest residue, ending the bar if leftover < 0.40m.
This ensures optimized usage of material and reduced waste.

🛠️ Tech Stack

C# WPF (XAML)

ObservableCollection + CollectionView

ClosedXML for Excel export

📊 Usage

Enter stock length.

Add required pieces.

Start cut → algorithm suggests best option.

Save & export results to Excel.
