import sys
import os
import subprocess
import xml.etree.ElementTree as ET
from PIL import Image
import shutil

from PySide6.QtCore import QObject, QObject, QUrl
from PySide6.QtWidgets import QApplication, QFileDialog, QWidget, QMessageBox
from PySide6.QtUiTools import QUiLoader
from PySide6.QtGui import QDesktopServices

from extract_pos_std import extract_gps_from_images

### IMPORTS ###
# pip install Pillow
# pip install PySide6
# pip install pyexiftool

class Interface(QObject):
    def __init__(self):
        super().__init__()
        ### LOAD UI ###
        self.ui = loader.load("mainwindow.ui", None)

        ### VARIABLES ###
        self.rep_chantier = ""
        self.images = []

        ### CONNECTIONS ###
        self.ui.action_ouvrir_chantier.triggered.connect(self.ouvre_chantier)
        self.ui.horizontalSlider_modele.valueChanged.connect(self.remplit_res_tapioca)
        self.ui.pushButton_construire_script.clicked.connect(self.construire_script)
        self.ui.pushButton_editer_script.clicked.connect(self.editer_script)
        self.ui.pushButton_executer_script.clicked.connect(self.execute_micmac)
        self.ui.checkBox_traitement_init.clicked.connect(self.active_traitement_init)
        self.ui.checkBox_mns_ortho.clicked.connect(self.active_mns_ortho)

        self.ui.action_aide_micmac.triggered.connect(self.ouvre_aide_micmac)
    
    def show(self):
        self.ui.show()

    def ouvre_aide_micmac(self):
        url = QUrl("https://micmac.ensg.eu/index.php/MicMac_tools")
        QDesktopServices.openUrl(url)

    def active_traitement_init(self, est_actif):
        self.ui.checkBox_oriconvert.setChecked(est_actif)
        self.ui.checkBox_tapioca.setChecked(est_actif)
        self.ui.checkBox_tapas.setChecked(est_actif)
        self.ui.checkBox_centerbascule.setChecked(est_actif)
        self.ui.checkBox_campari.setChecked(est_actif)
        self.ui.checkBox_apericloud.setChecked(est_actif)

    def active_mns_ortho(self, est_actif):
        self.ui.checkBox_malt.setChecked(est_actif)
        self.ui.checkBox_malt_ortho.setChecked(est_actif)

    def active_widgets(self):
        widgets = self.ui.findChildren(QWidget)
        for widget in widgets:
            widget.setEnabled(True)

    def charge_donnees_images(self, repertoire, images, purger=True):
        donnees_gps = os.path.join(repertoire, "CoordonneesGps.txt")
        incertitudes_gps = os.path.join(repertoire, "IncertitudesGps.txt")
        images_sans_gps = os.path.join(repertoire, "ImagesSansGps.txt")
        if purger:
            files = [donnees_gps, incertitudes_gps, images_sans_gps]
            for file in files:
                if os.path.exists(file):
                    os.remove(file)
        extract_gps_from_images(repertoire, images)
        if os.path.exists(images_sans_gps):
            with open(images_sans_gps, "r") as fichier:
                contenu_fichier = fichier.read()
                if contenu_fichier:
                    msg_box = QMessageBox()
                    msg_box.setIcon(QMessageBox.Critical)
                    msg_box.setWindowTitle("Erreur")
                    msg_box.setText("Certains fichiers sont dépourvus d'infos GPS.\nVérifier fichier NoGpsFromExif.txt")
                    msg_box.setDefaultButton(QMessageBox.Ok)
                    msg_box.exec()
                    return -1
        else:
            return 0
                
    def remplit_incertitudes_campari(self):
        incertitudes_gps = os.path.join(self.rep_chantier, "IncertitudesGps.txt")
        if os.path.exists(incertitudes_gps):
            with open(incertitudes_gps, 'r') as fichier:
                somme_inc_xy = 0
                somme_inc_z = 0
                compteur = 0
                for ligne in fichier:
                    colonnes = ligne.split()
                    if len(colonnes) == 4:
                        _, inc_x, inc_y, inc_z = colonnes
                        inc_x = float(inc_x)
                        inc_y = float(inc_y)
                        inc_z = float(inc_z)
                        inc_xy = ((inc_x**2)+(inc_y**2))**0.5
                        somme_inc_xy += inc_xy
                        somme_inc_z += inc_z
                        compteur += 1
                inc_xy_moy = somme_inc_xy/compteur
                inc_z_moy = somme_inc_z/compteur
                print(f"Incertitude XY moyenne : {inc_xy_moy:.2} mètres")
                print(f"Incertitude Z moyenne : {inc_z_moy:.2} mètres")
                self.ui.doubleSpinBox_sigma_XY.setValue(inc_xy_moy)
                self.ui.doubleSpinBox_sigma_Z.setValue(inc_z_moy)

    def verifie_presence_images(self, repertoire):
        compteur = 0
        images = []
        for fichier in os.listdir(repertoire):
            if fichier.lower().endswith("jpg"):
                compteur += 1
                images.append(fichier)
        return images, compteur

    def ouvre_chantier(self):
        options = QFileDialog.Options()
        repertoire = QFileDialog.getExistingDirectory(self.ui, "Sélectionner un dossier", "", options)
        if repertoire:
            images, nb_images = self.verifie_presence_images(repertoire)
            print(f"{nb_images} fichiers JPG dans {repertoire}")
            if nb_images > 0:
                if self.charge_donnees_images(repertoire, images) == 0:
                    self.images = images
                    self.rep_chantier = repertoire
                    self.ui.label_rep_chantier.setText(repertoire)
                    self.active_widgets()
                    self.remplit_res_tapioca(self.ui.horizontalSlider_modele.value())
                    self.remplit_incertitudes_campari()
            else:
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Critical)
                msg_box.setWindowTitle("Erreur")
                msg_box.setText("Pas de fichier JPG dans le dossier.\nChantier non ouvert.")
                msg_box.setDefaultButton(QMessageBox.Ok)
                msg_box.exec()

    def remplit_res_tapioca(self, valeur):
        tree = ET.parse("xml/presets.xml")
        root = tree.getroot()
        facteur_res_tapioca = float(root.find(f"./Tapioca/ResFactor/M{valeur}").text)
        image0 = os.path.join(self.rep_chantier, self.images[0])
        with Image.open(image0) as img:
            largeur = img.size[0]
            self.ui.spinBox_tapioca_res.setValue(int(largeur * facteur_res_tapioca))

    def construire_script(self):
        chemin_fichier_bat = os.path.join(self.rep_chantier, "execute_micmac.bat")
        with open(chemin_fichier_bat, "w") as fichier_bat:

            projection = self.ui.comboBox_proj.currentText()

            if self.ui.checkBox_oriconvert.isChecked():
                shutil.copy(f"projections/{projection}.xml", self.rep_chantier)
                fichier_bat.write(f"mm3d OriConvert \"#F=N_X_Y_Z\" CoordonneesGps.txt {projection}-Georef NameCple=Couples.xml ChSys=DegreeWGS84@{projection}.xml\n")

            if self.ui.checkBox_tapioca.isChecked():
                fichier_bat.write(f"mm3d Tapioca File Couples.xml {self.ui.spinBox_tapioca_res.value()}\n")
                fichier_bat.write(f"mm3d Schnaps \".*JPG\"\n")
            
            if self.ui.checkBox_tapas.isChecked():
                tapas_mode = self.ui.comboBox_tapas_mode.currentText()
                fichier_bat.write(f"mm3d Tapas {tapas_mode} \".*JPG\" SH=Homol_mini Out=Arbitraire\n")
            
            if self.ui.checkBox_centerbascule.isChecked():
                fichier_bat.write(f"mm3d CenterBascule \".*JPG\" Arbitraire {projection}-Georef {projection}-Brut\n")

            if self.ui.checkBox_campari.isChecked():
                fichier_bat.write(f"mm3d Campari \".*JPG\" {projection}-Brut {projection} EmGPS=[{projection}-Georef,{self.ui.doubleSpinBox_sigma_XY.value()},{self.ui.doubleSpinBox_sigma_Z.value()}] AllFree=1 GpsLa=[0,0,0]\n")

            if self.ui.checkBox_apericloud.isChecked():
                fichier_bat.write(f"mm3d AperiCloud \".*JPG\" Arbitraire\n")

            if self.ui.checkBox_malt.isChecked():
                mode_malt = self.ui.comboBox_malt_mode.currentText()
                ligne = f"mm3d Malt {mode_malt} \".*JPG\" {projection} EZA=1"
                if self.ui.checkBox_malt_ortho.isChecked():
                    ligne += " DoOrtho=1"
                suffixe_malt = self.ui.lineEdit_suffix_rep_malt.text()
                rep_mec = "MEC-Malt"
                rep_of = "Ortho-MEC-Malt"
                if suffixe_malt:
                    rep_mec += f"-{suffixe_malt}"
                    rep_of += f"-{suffixe_malt}"
                    ligne += f" DirMEC={rep_mec} DirOF={rep_of}"
                if self.ui.doubleSpinBox_resolterrain.value() != 0.0:
                    ligne += f" ResolTerrain={self.ui.doubleSpinBox_resolterrain.value()}"
                ligne += "\n"
                fichier_bat.write(ligne)
                fichier_bat.write(f"mm3d Tawny {rep_of}\n")
                chemin_tfw_src = os.path.join(self.rep_chantier, rep_of, "Orthophotomosaic.tfw")
                chemin_tfw_dest = os.path.join(self.rep_chantier, rep_of, "_Orthophoto.tfw")
                fichier_bat.write(f"mm3d ConvertIm {rep_of}/Orthophotomosaic.tif Out={rep_of}/_Orthophoto.tif\n")
                fichier_bat.write(f"copy \"{chemin_tfw_src}\" \"{chemin_tfw_dest}\"\n")

    def editer_script(self):
        chemin_fichier_bat = os.path.join(self.rep_chantier, "execute_micmac.bat")
        if os.path.exists(chemin_fichier_bat):
            try:
                subprocess.Popen(['notepad.exe', chemin_fichier_bat])
            except Exception as e:
                print(f"Failed to open {chemin_fichier_bat} with notepad.exe: {e}")

    def execute_micmac(self) :
        chemin_fichier_bat = os.path.join(self.rep_chantier, "execute_micmac.bat")
        subprocess.Popen(["start", "cmd", "/k", chemin_fichier_bat], shell=True, cwd=self.rep_chantier)


if __name__ == "__main__" :
    loader = QUiLoader()
    app = QApplication(sys.argv)
    window = Interface()

    window.show()
    app.exec()
