\* Essencial_LP *\
Maximize
OBJ: 0.000915817965978 Gás_de_295_para_GASDUC
 + 0.00121828581661 Gás_de_295_para_UPGN
 + 0.00111074640088 Gás_de_295_para_URGN
 + 0.000367387248141 Gás_de_295_para_URLs
 + 0.000916789210065 Gás_de_299_para_GASDUC
 + 0.0011958059861 Gás_de_299_para_UPGN
 + 0.00110932767928 Gás_de_299_para_URGN
 + 0.00036345278084 Gás_de_299_para_URLs
 + 0.00109203771603 Gás_de_302_para_MIX
 + 0.00122538860689 Gás_de_302_para_UPGN
 + 0.000367943636624 Gás_de_302_para_URLs
Subject To
Alocar_todo_o_gás_do_295_e_do_Reciclo: Gás_de_295_para_GASDUC
 + Gás_de_295_para_UPGN + Gás_de_295_para_URGN + Gás_de_295_para_URLs
 = 1875215.01404
Alocar_todo_o_gás_do_299: Gás_de_299_para_GASDUC + Gás_de_299_para_UPGN
 + Gás_de_299_para_URGN + Gás_de_299_para_URLs = 3240281.2056
Alocar_todo_o_gás_do_302: Gás_de_302_para_MIX + Gás_de_302_para_UPGN
 + Gás_de_302_para_URLs = 16784566.0723
GASDUC_máximo: Gás_de_295_para_GASDUC + Gás_de_299_para_GASDUC <= 2000000
GASDUC_mínimo: - Gás_de_295_para_GASDUC - Gás_de_299_para_GASDUC <= 0
GV_C1: - 4.46245781308e-07 Gás_de_302_para_MIX >= -1.9938898051
MIX_máximo: Gás_de_302_para_MIX <= 8000000
MIX_mínimo: - Gás_de_302_para_MIX <= -1e-06
UPGN_máximo: Gás_de_295_para_UPGN + Gás_de_299_para_UPGN
 + Gás_de_302_para_UPGN <= 5400000
UPGN_mínimo: - Gás_de_295_para_UPGN - Gás_de_299_para_UPGN
 - Gás_de_302_para_UPGN <= -1e-06
URGN_máximo: Gás_de_295_para_URGN + Gás_de_299_para_URGN <= 3000000
URGN_mínimo: - Gás_de_295_para_URGN - Gás_de_299_para_URGN <= -1e-06
URLs_máximo: Gás_de_295_para_URLs + Gás_de_299_para_URLs
 + Gás_de_302_para_URLs <= 16200000
URLs_mínimo: - Gás_de_295_para_URLs - Gás_de_299_para_URLs
 - Gás_de_302_para_URLs <= -8100000
Bounds
 1e-06 <= Gás_de_295_para_GASDUC <= 2000000
 1e-06 <= Gás_de_295_para_UPGN <= 5400000
 1e-06 <= Gás_de_295_para_URGN <= 3000000
 1e-06 <= Gás_de_295_para_URLs <= 16200000
 1e-06 <= Gás_de_299_para_GASDUC <= 2000000
 Gás_de_299_para_UPGN <= 5400000
 1e-06 <= Gás_de_299_para_URGN <= 3000000
 1e-06 <= Gás_de_299_para_URLs <= 16200000
 1e-06 <= Gás_de_302_para_MIX <= 8000000
 1e-06 <= Gás_de_302_para_UPGN <= 5400000
 1e-06 <= Gás_de_302_para_URLs <= 16200000
End
