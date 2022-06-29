dimm_types_dict = {"RU": {1: "Общий",
                          2: "Питтинг",
                          3: "Язва",
                          4: "Продольная канавка",
                          5: "Продольная риска",
                          6: "Поперечная канавка",
                          7: "Поперечная риска",
                          },
                   "EN": {1: "General",
                          2: "Pitting",
                          3: "Pinhole",
                          4: "Axial Grooving",
                          5: "Axial Slotting",
                          6: "Circumferential Grooving",
                          7: "Circumferential Slotting",
                          },
                   "SHORT": {1: "GENE",
                             2: "PITT",
                             3: "PINH",
                             4: "AXGR",
                             5: "AXSL",
                             6: "CIGR",
                             7: "CISL",
                             }
                   }


def dimm(length, width, wt, return_format="RU"):
    if return_format not in ["RU", "EN", "SHORT"]:
        return_format = "RU"

    dimm_dict = dimm_types_dict[return_format]

    if wt < 10:
        wt = 10

    if length == "" or width == "" or wt == "":
        return "some is empty"
    if width >= 3 * wt and length >= 3 * wt:
        return dimm_dict[1]
    if ((1 * wt <= width < 6 * wt) and (1 * wt <= length < 6 * wt)) and (
             0.5 <= length / width < 2) and (width < 3 * wt or length < 3 * wt):
        return dimm_dict[2]
    if (0 < width < 1 * wt) and (0 < length < 1 * wt):
        return dimm_dict[3]
    if (1 * wt <= width < 3 * wt) and (length / width >= 2):
        return dimm_dict[4]
    if (0 < width < 1 * wt) and (length >= 1 * wt):
        return dimm_dict[5]
    if (1 * wt <= length < 3 * wt) and (length / width <= 0.5):
        return dimm_dict[6]
    if (0 < length < 1 * wt) and (width >= 1 * wt):
        return dimm_dict[7]


if __name__ == '__main__':
    print(dimm(21, 40, 6.5, "RU"))
