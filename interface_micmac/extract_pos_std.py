import os
import csv
import exiftool

def extract_gps_from_images(directory, images):
    gps_data = []
    gps_std_data = []
    no_gps_data = []

    output_gps_data = os.path.join(directory, "CoordonneesGps.txt")
    output_gps_std_data = os.path.join(directory, "IncertitudesGps.txt")
    output_no_gps_data = os.path.join(directory, "ImagesSansGps.txt")

    with exiftool.ExifToolHelper() as et:
        for image in images:
            image_path = os.path.join(directory, image)

            try:
                metadata_list = et.get_metadata(image_path)
                if not metadata_list:
                    print(f"Pas de métadonnées trouvée dans : {image}")
                    no_gps_data.append([image])
                    continue

                metadata = metadata_list[0]

                if 'XMP:GPSLongitude' in metadata.keys() and 'XMP:GPSLatitude' in metadata.keys() and 'XMP:AbsoluteAltitude' in metadata.keys():
                    if 'XMP:RtkStdLon' in metadata.keys() and 'XMP:RtkStdLat' in metadata.keys() and 'XMP:RtkStdHgt' in metadata.keys():
                        gps_data.append([
                            image,
                            float(metadata['XMP:GPSLongitude']),
                            float(metadata['XMP:GPSLatitude']),
                            float(metadata['XMP:AbsoluteAltitude'])
                        ])
                                            
                        gps_std_data.append([
                            image,
                            float(metadata['XMP:RtkStdLon']),
                            float(metadata['XMP:RtkStdLat']),
                            float(metadata['XMP:RtkStdHgt'])
                        ])

                    print(f"Données GPS ajoutées depuis {image}")
                else:
                    no_gps_data.append([image])
                    print(f"Pas de données GPS dans {image}")

            except exiftool.exceptions.ExifToolExecuteError as e:
                print(f"ExifTool execution error for image {image}: {e}")
                no_gps_data.append([image])
            except Exception as e:
                print(f"Unexpected error for image {image}: {e}")
                no_gps_data.append([image])

    if gps_data:
        with open(output_gps_data, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter=' ')
            writer.writerows(gps_data)

    if gps_std_data:
        with open(output_gps_std_data, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter=' ')
            writer.writerows(gps_std_data)

    if no_gps_data:
        with open(output_no_gps_data, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter=' ')
            writer.writerows(no_gps_data)

if __name__ == '__main__':
    directory = "test"
    images = os.listdir(directory)
    extract_gps_from_images(directory, images)
