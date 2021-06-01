from pathlib import Path

source_path = Path("fileGenerated/abde.text")
destination_path = Path("fileGenerated/abdelo.text")
with source_path.open("rb") as source:
    with destination_path.open("wb") as destination:
        bytes = source.read(5)
        while len(bytes) > 0:
           print(f"{bytes} => {bytes[:4]}")
           destination.write(bytes[:4])
           bytes = source.read(5)

destination_path.rename(source_path)



