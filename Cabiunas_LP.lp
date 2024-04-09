\* Cabiunas_LP *\
Maximize
OBJ: 0.000924811673081 Gás_de_295_para_GASDUC
 + 0.001230190417 Gás_de_295_para_UPGN
 + 1.02037192846e-09 Gás_de_295_para_URGN
 + 0.00113209629472 Gás_de_295_para_URLs
 + 0.000884627452647 Gás_de_299_para_GASDUC
 + 0.0011968667161 Gás_de_299_para_UPGN
 - 4.00394706994e-05 Gás_de_299_para_URGN
 + 0.00109487456267 Gás_de_299_para_URLs
 + 0.00109286495293 Gás_de_302_para_MIX
 + 0.00122709346557 Gás_de_302_para_UPGN
 + 0.00111725104996 Gás_de_302_para_URLs
Subject To
Alocar_todo_o_gás_do_295_e_do_Reciclo: Gás_de_295_para_GASDUC
 + Gás_de_295_para_UPGN + Gás_de_295_para_URGN + Gás_de_295_para_URLs
 = 1647992.36148
Alocar_todo_o_gás_do_299: Gás_de_299_para_GASDUC + Gás_de_299_para_UPGN
 + Gás_de_299_para_URGN + Gás_de_299_para_URLs = 3240281.2056
Alocar_todo_o_gás_do_302: Gás_de_302_para_MIX + Gás_de_302_para_UPGN
 + Gás_de_302_para_URLs = 16784566.0723
GASDUC_máximo: Gás_de_295_para_GASDUC + Gás_de_299_para_GASDUC <= 2000000
GASDUC_mínimo: - Gás_de_295_para_GASDUC - Gás_de_299_para_GASDUC <= 0
MIX_máximo: Gás_de_302_para_MIX <= 16000000
MIX_mínimo: - Gás_de_302_para_MIX <= -1e-06
UPGN_máximo: Gás_de_295_para_UPGN + Gás_de_299_para_UPGN
 + Gás_de_302_para_UPGN <= 5400000
UPGN_mínimo: - Gás_de_295_para_UPGN - Gás_de_299_para_UPGN
 - Gás_de_302_para_UPGN <= -100
URGN_máximo: Gás_de_295_para_URGN + Gás_de_299_para_URGN <= 300
URGN_mínimo: - Gás_de_295_para_URGN - Gás_de_299_para_URGN <= -1e-06
URLs_máximo: Gás_de_295_para_URLs + Gás_de_299_para_URLs
 + Gás_de_302_para_URLs <= 16200000
URLs_mínimo: - Gás_de_295_para_URLs - Gás_de_299_para_URLs
 - Gás_de_302_para_URLs <= -8100000
Bounds
 1e-06 <= Gás_de_295_para_GASDUC <= 2000000
 1e-06 <= Gás_de_295_para_UPGN <= 5400000
 1e-06 <= Gás_de_295_para_URGN <= 300
 1e-06 <= Gás_de_295_para_URLs <= 16200000
 1e-06 <= Gás_de_299_para_GASDUC <= 2000000
 Gás_de_299_para_UPGN <= 5400000
 1e-06 <= Gás_de_299_para_URGN <= 300
 1e-06 <= Gás_de_299_para_URLs <= 16200000
 1e-06 <= Gás_de_302_para_MIX <= 16000000
 1e-06 <= Gás_de_302_para_UPGN <= 5400000
 1e-06 <= Gás_de_302_para_URLs <= 16200000
End
