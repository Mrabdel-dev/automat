import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QMessageBox


def dialog():
    mbox = QMessageBox()

    mbox.setText("Your allegiance has been noted")
    mbox.setDetailedText("You are now a disciple and subject of the all-knowing Guru")
    mbox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

    mbox.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = QWidget()
    w.resize(500, 500)
    w.setWindowTitle('ManeoReseaux')

    label = QLabel(w)
    label.setText("button wich start pop up window")
    label.move(150, 170)
    label.show()

    btn = QPushButton(w)
    btn.setText('start')
    btn.move(170, 200)
    btn.show()
    btn.clicked.connect(dialog)

    w.show()
    sys.exit(app.exec_())