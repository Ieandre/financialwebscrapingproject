{
  "nbformat": 4,
  "nbformat_minor": 0,
  "metadata": {
    "colab": {
      "name": "WebScrappingProject",
      "provenance": [],
      "collapsed_sections": [],
      "authorship_tag": "ABX9TyN5t7IiPGmG51qnMvxwMVq2"
    },
    "kernelspec": {
      "name": "python3",
      "display_name": "Python 3"
    },
    "language_info": {
      "name": "python"
    }
  },
  "cells": [
    {
      "cell_type": "markdown",
      "source": [
        "# Installation des librairies demandées"
      ],
      "metadata": {
        "id": "WkR-Q88plBWI"
      }
    },
    {
      "cell_type": "code",
      "execution_count": 4,
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/",
          "height": 131
        },
        "id": "LnG33uxIk6Qp",
        "outputId": "72bc8b21-8a83-4ef0-9ef7-89368ad145ec"
      },
      "outputs": [
        {
          "output_type": "error",
          "ename": "SyntaxError",
          "evalue": "ignored",
          "traceback": [
            "\u001b[0;36m  File \u001b[0;32m\"<ipython-input-4-5f6c4c3c43d3>\"\u001b[0;36m, line \u001b[0;32m2\u001b[0m\n\u001b[0;31m    pip install openpyxl\u001b[0m\n\u001b[0m              ^\u001b[0m\n\u001b[0;31mSyntaxError\u001b[0m\u001b[0;31m:\u001b[0m invalid syntax\n"
          ]
        }
      ],
      "source": [
        "#pip install beautifulsoup4\n"
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "import urllib.request\n",
        "import openpyxl\n",
        "from bs4 import BeautifulSoup\n",
        "\n",
        "wb = openpyxl.load_workbook(r'test.xlsx')\n",
        "sheet = wb.active\n",
        "valeurlien = sheet['F121'].value\n",
        "print(valeurlien)\n",
        "\n",
        "def Convert(string):\n",
        "    li = list(string.split(\" \"))\n",
        "    return li\n",
        "\n",
        "\n",
        "soup = BeautifulSoup(urllib.request.urlopen(valeurlien), 'lxml')\n",
        "\n",
        "\n",
        "tableau1 = soup('table', {\"class\" : \"BordCollapseYear2\"})[0]\n",
        "tableau2 = soup('table', {\"class\" : \"BordCollapseYear2\"})[1]\n",
        "\n",
        "\n",
        "  \n",
        "\n",
        "\n",
        "capitalisationtemp = Convert(tableau1.findAll('tr')[5].get_text(\" \"))\n",
        "capitalisation = capitalisationtemp[5:12]\n",
        "print(capitalisation)\n",
        "\n",
        "\n",
        "\n",
        "pertemp = Convert(tableau1.findAll('tr')[3].get_text(\" \"))\n",
        "per = pertemp[3:10]\n",
        "print(per)\n",
        "\n",
        "\n",
        "\n",
        "bnatemp = Convert(tableau2.findAll('tr')[8].get_text(\" \"))\n",
        "bna = bnatemp[4:11]\n",
        "print(bna)\n",
        "\n",
        "\n",
        "\n",
        "\n",
        "\n",
        "\n"
      ],
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "id": "KJMeuAiWldwH",
        "outputId": "c4f3274e-0b06-4933-a1f5-9d8f581779ce"
      },
      "execution_count": 30,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "http://www.zonebourse.com/ADOCIA-9894600/fondamentaux/\n",
            "['13,7x', '3,64x', '2,11x', '8,44x', '8,44x', '2,59x', '0,99x']\n",
            "['-53,0x', '-12,0x', '16,5x', '-3,67x', '-2,52x', '-2,33x', '3,98x']\n",
            "['-1,15', '-1,20', '1,00', '-2,70', '-3,30', '-3,01', '1,76']\n"
          ]
        }
      ]
    },
    {
      "cell_type": "code",
      "source": [
        "import urllib.request\n",
        "import openpyxl\n",
        "from bs4 import BeautifulSoup\n",
        "\n",
        "soup = BeautifulSoup(urllib.request.urlopen(valeurlien), 'lxml')\n",
        "\n",
        "\n",
        "wb = openpyxl.load_workbook(r'test.xlsx')\n",
        "sheet = wb.active\n",
        "valeurlien = sheet['F120'].value\n",
        "print(valeurlien)\n",
        "\n",
        "def Convert(string):\n",
        "    li = list(string.split(\" \"))\n",
        "    return li\n",
        "\n",
        "for cell in sheet[\"F\"]:\n",
        "  if type(cell.value) is str:\n",
        "    if cell.value != \"Link\":\n",
        "      print(cell.value)\n",
        "      soup = BeautifulSoup(urllib.request.urlopen(cell.value), 'lxml')\n",
        "      tableau1 = soup('table', {\"class\" : \"BordCollapseYear2\"})[0]\n",
        "      tableau2 = soup('table', {\"class\" : \"BordCollapseYear2\"})[1]\n",
        "\n",
        "      capitalisationtemp = Convert(tableau1.findAll('tr')[5].get_text(\" \"))\n",
        "      capitalisation = capitalisationtemp[5:12]\n",
        "      print(capitalisation)\n",
        "\n",
        "      pertemp = Convert(tableau1.findAll('tr')[3].get_text(\" \"))\n",
        "      per = pertemp[3:10]\n",
        "      print(per)\n",
        "\n",
        "      bnatemp = Convert(tableau2.findAll('tr')[8].get_text(\" \"))\n",
        "      bna = bnatemp[4:11]\n",
        "      print(bna)\n"
      ],
      "metadata": {
        "colab": {
          "base_uri": "https://localhost:8080/"
        },
        "id": "nhBaz6cv6dz5",
        "outputId": "f810f9a5-2bbb-47da-d2b6-231ef9460bc3"
      },
      "execution_count": 37,
      "outputs": [
        {
          "output_type": "stream",
          "name": "stdout",
          "text": [
            "http://www.zonebourse.com/AB-SCIENCE-6133795/fondamentaux/\n",
            "http://www.zonebourse.com/AB-SCIENCE-6133795/fondamentaux/\n",
            "['182x', '345x', '199x', '83,7x', '151x', '566x', '\\n']\n",
            "['-15,7x', '-17,6x', '-11,1x', '-5,03x', '-9,75x', '-57,8x', '\\n']\n",
            "['-0,78', '-0,78', '-0,75', '-0,69', '-0,55', '-0,34', '\\n']\n",
            "http://www.zonebourse.com/ADOCIA-9894600/fondamentaux/\n",
            "['13,7x', '3,64x', '2,11x', '8,44x', '8,44x', '2,59x', '0,99x']\n",
            "['-53,0x', '-12,0x', '16,5x', '-3,67x', '-2,52x', '-2,33x', '3,98x']\n",
            "['-1,15', '-1,20', '1,00', '-2,70', '-3,30', '-3,01', '1,76']\n",
            "http://www.zonebourse.com/BASTIDE-LE-CONFORT-MED-5023/fondamentaux/\n",
            "['1,19x', '1,25x', '0,84x', '0,66x', '0,78x', '0,70x', '0,67x']\n",
            "['48,1x', '50,4x', '62,1x', '22,0x', '26,0x', '17,7x', '16,0x']\n",
            "['0,74', '0,98', '0,62', '1,58', '1,83', '2,53', '2,80']\n"
          ]
        }
      ]
    }
  ]
}