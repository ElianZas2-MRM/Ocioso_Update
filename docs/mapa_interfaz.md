# Mapa de `main_interface.py`

El archivo son **5758 lineas**, y 4944 de ellas viven dentro de
**una sola funcion**: `iniciar_interfaz()`. No esta partida, asi que este mapa es la forma
de moverse por ella sin leerla entera.

> Generado midiendo el archivo. Si se mueve codigo, los numeros de linea cambian:
> regeneralo en vez de corregirlo a mano.

## Fuera de la funcion grande

| Que | Lineas |
|---|---|
| `_excel_path_for` (funcion) | 79-87 |
| `_t3_tag` (funcion) | 90-92 |
| `_etiqueta_t3` (funcion) | 100-102 |
| `_lead_excel_name` (funcion) | 105-107 |
| `_generic_excel_path_for` (funcion) | 110-112 |
| `_device_excel_suffix` (funcion) | 115-123 |
| `_ensure_serialized_setup` (funcion) | 132-146 |
| `_DemoSchedulerPanel` (clase) | 220-629 |
| `get_button_icon` (funcion) | 637-650 |
| `create_demo_interface` (funcion) | 5754-5755 |

## Dentro de `iniciar_interfaz()`

Las funciones anidadas, en el orden en que aparecen. La columna **captura** dice cuantos
nombres del closure usa cada una: es la medida de que tan enredada esta con el resto, y
por lo tanto que tan facil seria sacarla afuera.

| Funcion | Lineas | Largo | Captura |
|---|---|---|---|
| `_forzar_foreground` | 826-858 | 32 | 2 |
| `_dark_titlebar` | 888-897 | 9 | 1 |
| `make_icon_btn` | 904-957 | 53 | 7 |
| `log_message` | 1007-1008 | 1 | 0 |
| `_render_fixed_header` | 1040-1052 | 12 | 6 |
| `_actualizar_drivers` | 1070-1083 | 13 | 5 |
| `_on_pause_change` | 1110-1115 | 5 | 3 |
| `_on_visible_change` | 1117-1119 | 2 | 2 |
| `toggle_email_options` | 1189-1195 | 6 | 3 |
| `switch_tab` | 1218-1229 | 11 | 7 |
| `make_scrollable_tab_container` | 1264-1283 | 19 | 5 |
| `_abrir_ids_dinamicos` | 1291-1301 | 10 | 5 |
| `make_pill_group` | 1350-1405 | 55 | 9 |
| `_load_lt_creds` | 1442-1460 | 18 | 2 |
| `toggle_disp` | 1462-1492 | 30 | 16 |
| `_save_ui_prefs` | 1525-1570 | 45 | 19 |
| `_refresh_t3_also_state` | 1618-1626 | 8 | 4 |
| `refresh_ver_nav_state` | 1629-1633 | 4 | 2 |
| `_refresh_excel_par_warning` | 1651-1663 | 12 | 3 |
| `_save_lt_creds` | 1688-1699 | 11 | 7 |
| `actualizar_warning` | 1720-1727 | 7 | 5 |
| `abrir_config_avanzada` | 1729-1817 | 88 | 19 |
| `refresh_all_country_ui` | 1830-1863 | 33 | 10 |
| `toggle_country` | 1865-1867 | 2 | 2 |
| `select_all_countries` | 1869-1872 | 3 | 3 |
| `select_none_countries` | 1874-1877 | 3 | 3 |
| `build_country_selection_card` | 1879-1931 | 52 | 15 |
| `_cfg_for` | 1965-1966 | 1 | 2 |
| `_state_for` | 1968-1969 | 1 | 2 |
| `get_current_cfg` | 1971-1972 | 1 | 2 |
| `_sched_summary` | 1974-1979 | 5 | 3 |
| `set_execution_mode` | 1981-1986 | 5 | 2 |
| `select_sched_mode` | 1988-1990 | 2 | 2 |
| `toggle_sched_nav` | 1992-2002 | 10 | 3 |
| `switch_sched_subtab` | 2016-2029 | 13 | 10 |
| `_build_sched_block` | 2048-2175 | 127 | 32 |
| `_mk_sched_badge` | 2177-2181 | 4 | 2 |
| `_refresh_prog` | 2183-2246 | 63 | 21 |
| `_refresh_sched_config_ui` | 2248-2249 | 1 | 1 |
| `_on_sched_save` | 2258-2267 | 9 | 4 |
| `open_scheduler_config_popup` | 2269-2303 | 34 | 11 |
| `toggle_scheduler_config` | 2305-2306 | 1 | 1 |
| `_sched_activate` | 2308-2340 | 32 | 11 |
| `_sched_activate_masivo` | 2342-2380 | 38 | 18 |
| `_sched_deactivate` | 2382-2390 | 8 | 5 |
| `_sched_deactivate_masivo` | 2392-2400 | 8 | 5 |
| `_refresh_deact_btn` | 2402-2425 | 23 | 6 |
| `_sched_run_now` | 2427-2429 | 2 | 2 |
| `_horas_programadas_leads` | 2523-2527 | 4 | 1 |
| `_refresh_win_task_lbl` | 2529-2540 | 11 | 5 |
| `cmd_registrar_win_task` | 2545-2574 | 29 | 8 |
| `cmd_quitar_win_task` | 2576-2587 | 11 | 5 |
| `_on_toggle_abrir_app` | 2607-2626 | 19 | 9 |
| `view_results_dialog_leads` | 2635-2642 | 7 | 5 |
| `_sched_load_triggered` | 2657-2664 | 7 | 3 |
| `_sched_save_triggered` | 2666-2676 | 10 | 3 |
| `_sched_marcar_ventana_actual` | 2678-2697 | 19 | 6 |
| `_sched_monitor` | 2699-2749 | 50 | 13 |
| `select_p_tab` | 2768-2778 | 10 | 8 |
| `_available_device_excels` | 2802-2812 | 10 | 3 |
| `_on_preview_device_change` | 2820-2824 | 4 | 4 |
| `show_table_msg` | 2837-2855 | 18 | 3 |
| `cmd_agregar` | 2858-2868 | 10 | 6 |
| `cmd_eliminar` | 2870-2880 | 10 | 7 |
| `cmd_clonar` | 2882-2895 | 13 | 7 |
| `cmd_actualizar_tbl` | 2897-2900 | 3 | 5 |
| `cmd_guardar_tbl` | 2902-2917 | 15 | 13 |
| `cmd_abrir_excel` | 2919-2931 | 12 | 12 |
| `_close_cell_editor` | 2987-2994 | 7 | 1 |
| `_edit_cell` | 2996-3038 | 42 | 8 |
| `_stripe_rows` | 3048-3050 | 2 | 1 |
| `_refresh_tab_count` | 3052-3055 | 3 | 1 |
| `update_table_data` | 3058-3125 | 67 | 15 |
| `_build_retry_excel` | 3156-3189 | 33 | 2 |
| `_eligible_retry_targets` | 3191-3209 | 18 | 0 |
| `execute_send_leads` | 3212-4266 | 1054 | 82 |
| `view_results_dialog` | 4268-4275 | 7 | 5 |
| `_retry_failed_leads` | 4296-4323 | 27 | 7 |
| `refresh_execute_state` | 4333-4337 | 4 | 5 |
| `buscar_excel_masivo` | 4441-4445 | 4 | 3 |
| `select_sync_report` | 4511-4515 | 4 | 3 |
| `run_sync_action` | 4528-4656 | 128 | 15 |
| `cmd_iniciar_masivo` | 4706-5036 | 330 | 43 |
| `abrir_carpeta_masivo` | 5074-5080 | 6 | 4 |
| `_detect_excel_country` | 5116-5121 | 5 | 1 |
| `select_excel_pais` | 5138-5150 | 12 | 10 |
| `switch_url_mode` | 5186-5199 | 13 | 8 |
| `refresh_doc_types_section` | 5279-5292 | 13 | 9 |
| `_selected_doc_types` | 5294-5299 | 5 | 4 |
| `update_excel_calculation` | 5312-5353 | 41 | 10 |
| `toggle_excel_disp` | 5355-5363 | 8 | 7 |
| `_on_excel_url_change` | 5384-5392 | 8 | 6 |
| `_build_excel_pares` | 5400-5409 | 9 | 2 |
| `_email_device_token` | 5411-5414 | 3 | 1 |
| `_rows_for_pais` | 5416-5437 | 21 | 1 |
| `_do_generar` | 5439-5498 | 59 | 23 |
| `cmd_generar_excels` | 5500-5501 | 1 | 1 |
| `cmd_regen_datos` | 5503-5504 | 1 | 1 |
| `cmd_borrar_urls` | 5506-5511 | 5 | 4 |
| `_on_global_mousewheel` | 5548-5572 | 24 | 2 |
| `_restaurar_mousewheel` | 5575-5580 | 5 | 2 |
| `_restore_from_tray` | 5594-5603 | 9 | 2 |
| `_force_close` | 5605-5622 | 17 | 3 |
| `_menu_tray` | 5624-5678 | 54 | 5 |
| `_bombear_eventos_tray` | 5680-5693 | 13 | 3 |
| `_on_close` | 5695-5721 | 26 | 10 |

## Las mas grandes, esten al nivel que esten

| Funcion | Lineas | Largo | Captura |
|---|---|---|---|
| `execute_send_leads` | 3212-4266 | 1054 | 82 |
| `cmd_iniciar_masivo` | 4706-5036 | 330 | 43 |
| `run_sync_action` | 4528-4656 | 128 | 15 |
| `_build_sched_block` | 2048-2175 | 127 | 32 |
| `show_completed` | 3898-4025 | 127 | 18 |
| `sync_worker` | 4555-4653 | 98 | 11 |
| `_run_session` | 4115-4208 | 93 | 31 |
| `abrir_config_avanzada` | 1729-1817 | 88 | 19 |
| `worker` | 4960-5034 | 74 | 26 |
| `update_table_data` | 3058-3125 | 67 | 15 |
| `_refresh_prog` | 2183-2246 | 63 | 21 |
| `_do_generar` | 5439-5498 | 59 | 23 |

## Si alguna vez se parte

Hay **176 funciones anidadas** compartiendo el closure. El dato que importa es que
cada una captura poco: las mas grandes usan entre 15 y 31 nombres del entorno, no cientos.
Las que capturan **cero** (`make_pill_group`, `make_icon_btn`) son helpers puros de UI y se
pueden sacar afuera sin tocar nada mas.

El orden razonable seria: primero las de captura 0, despues las mas grandes pasandoles
explicitamente lo que usan. Pero antes hace falta cobertura: hoy la UI **no tiene un solo
test**, asi que una cirugia asi seria a ciegas.
