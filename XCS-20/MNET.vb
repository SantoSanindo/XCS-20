Module MNET
    'Ring Status
    Declare Function B_mnet_enable_soft_watchdog Lib "MotionNet.dll" Alias "_mnet_enable_soft_watchdog" (ByVal RingNo As Integer, UserEvent As Long) As Integer
    Declare Function B_mnet_set_ring_quality_param Lib "MotionNet.dll" Alias "_mnet_set_ring_quality_param" (ByVal RingNo As Integer, ByVal ContinueErr As Integer, ByVal ErrorRate As Integer) As Integer

    'Ring Operation
    Declare Function B_mnet_reset_ring Lib "MotionNet.dll" Alias "_mnet_reset_ring" (ByVal RingNo As Integer) As Integer
    Declare Function B_mnet_get_ring_active_table Lib "MotionNet.dll" Alias "_mnet_get_ring_active_table" (ByVal RingNo As Integer, ActTable As Double) As Integer
    Declare Function B_mnet_get_slave_info Lib "MotionNet.dll" Alias "_mnet_get_slave_info" (ByVal RingNo As Integer, ByVal RingNo As Integer) As Integer
    Declare Function B_mnet_start_ring Lib "MotionNet.dll" Alias "_mnet_start_ring" (ByVal RingNo As Integer) As Integer
    Declare Function B_mnet_stop_ring Lib "MotionNet.dll" Alias "_mnet_stop_ring" (ByVal RingNo As Integer) As Integer

    'Io Operation
    Declare Function B_mnet_io_output Lib "MotionNet.dll" Alias "_mnet_io_output" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal PortNo As Integer, ByVal Value As Integer) As Integer
    Declare Function B_mnet_io_input Lib "MotionNet.dll" Alias "_mnet_io_input" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal PortNo As Integer) As Integer

    'Axis Slave Operation
    Declare Function B_mnet_m1_initial Lib "MotionNet.dll" Alias "_mnet_m1_initial" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer

    ''Pulse In/Out
    Declare Function B_mnet_m1_set_pls_outmode Lib "MotionNet.dll" Alias "_mnet_m1_set_pls_outmode" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal pls_outmode As Integer) As Integer
    Declare Function B_mnet_m1_set_pls_iptmode Lib "MotionNet.dll" Alias "_mnet_m1_set_pls_iptmode" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal pls_iptmode As Integer, ByVal pls_logic As Integer) As Integer
    Declare Function B_mnet_m1_set_feedback_src Lib "MotionNet.dll" Alias "_mnet_m1_set_feedback_src" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal Src As Integer) As Integer


    ''Motion Interface I/O
    Declare Function B_mnet_m1_set_alm Lib "MotionNet.dll" Alias "_mnet_m1_set_alm" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal alm_logic As Integer, ByVal alm_mode As Integer) As Integer
    Declare Function B_mnet_m1_set_inp Lib "MotionNet.dll" Alias "_mnet_m1_set_inp" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal inp_enable As Integer, ByVal inp_logic As Integer) As Integer
    Declare Function B_mnet_m1_set_erc Lib "MotionNet.dll" Alias "_mnet_m1_set_erc" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal erc_logic As Integer, ByVal erc_on_time As Integer, ByVal erc_off_time As Integer) As Integer
    Declare Function B_mnet_m1_set_erc_on Lib "MotionNet.dll" Alias "_mnet_m1_set_erc_on" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal on_off As Integer) As Integer
    Declare Function B_mnet_m1_set_ralm Lib "MotionNet.dll" Alias "_mnet_m1_set_ralm" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal on_off As Integer) As Integer
    Declare Function B_mnet_m1_set_sd Lib "MotionNet.dll" Alias "_mnet_m1_set_sd" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal enable As Integer, ByVal sd_logic As Integer, ByVal sd_latch As Integer, ByVal sd_mode As Integer) As Integer
    Declare Function B_mnet_m1_set_svon Lib "MotionNet.dll" Alias "_mnet_m1_set_svon" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal on_off As Integer) As Integer
    Declare Function B_mnet_m1_set_pcs Lib "MotionNet.dll" Alias "_mnet_m1_set_pcs" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal pcs_logic As Integer) As Integer
    Declare Function B_mnet_m1_set_clr Lib "MotionNet.dll" Alias "_mnet_m1_set_clr" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal clr_logic As Integer) As Integer

    Declare Function B_mnet_m1_dio_output Lib "MotionNet.dll" Alias "_mnet_m1_dio_output" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal DoNO As Integer, ByVal on_off As Integer) As Integer
    Declare Function B_mnet_m1_dio_input Lib "MotionNet.dll" Alias "_mnet_m1_dio_input" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal DoNO As Integer) As Integer

    ''Stop
    Declare Function B_mnet_m1_sd_stop Lib "MotionNet.dll" Alias "_mnet_m1_sd_stop" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer
    Declare Function B_mnet_m1_emg_stop Lib "MotionNet.dll" Alias "_mnet_m1_emg_stop" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer

    ''IO Monitor
    Declare Function B_mnet_m1_get_io_status Lib "MotionNet.dll" Alias "_mnet_m1_get_io_status" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, io_sts As Integer) As Integer

    '' Motion Done
    Declare Function B_mnet_m1_motion_done Lib "MotionNet.dll" Alias "_mnet_m1_motion_done" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, motdone As Integer) As Integer

    '' Single Axis Motion
    Declare Function B_mnet_m1_set_tmove_speed Lib "MotionNet.dll" Alias "_mnet_m1_set_tmove_speed" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal StrVel As Double, ByVal MaxVel As Double, ByVal Tacc As Double, ByVal Tdec As Double) As Integer
    Declare Function B_mnet_m1_set_smove_speed Lib "MotionNet.dll" Alias "_mnet_m1_set_smove_speed" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal StrVel As Double, ByVal MaxVel As Double, ByVal Tacc As Double, ByVal Tdec As Double, ByVal SVacc As Double, ByVal SVdec As Double) As Integer
    Declare Function B_mnet_m1_v_change Lib "MotionNet.dll" Alias "_mnet_m1_v_change" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal NewVel As Double, ByVal TimeSecond As Double) As Integer
    Declare Function B_mnet_m1_fix_speed_range Lib "MotionNet.dll" Alias "_mnet_m1_fix_speed_range" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal MaxVel As Double) As Integer
    Declare Function B_mnet_m1_unfix_speed_range Lib "MotionNet.dll" Alias "_mnet_m1_unfix_speed_range" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer
    Declare Function B_mnet_m1_set_move_ratio Lib "MotionNet.dll" Alias "_mnet_m1_set_move_ratio" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal MoveRatio As Double) As Integer

    Declare Function B_mnet_m1_v_move Lib "MotionNet.dll" Alias "_mnet_m1_v_move" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal DIR As Integer) As Integer
    Declare Function B_mnet_m1_start_r_move Lib "MotionNet.dll" Alias "_mnet_m1_start_r_move" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal Distance As Long) As Integer
    Declare Function B_mnet_m1_start_a_move Lib "MotionNet.dll" Alias "_mnet_m1_start_a_move" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal Position As Long) As Integer


    ''Position Compare and latch
    Declare Function B_mnet_m1_set_comparator_mode Lib "MotionNet.dll" Alias "_mnet_m1_set_comparator_mode" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal CmpNo As Integer, ByVal CmpSrc As Integer, ByVal CmpMethod As Integer, ByVal CmpAction As Integer) As Integer
    Declare Function B_mnet_m1_set_comparator_data Lib "MotionNet.dll" Alias "_mnet_m1_set_comparator_data" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal CmpNo As Integer, ByVal Data As Double) As Integer
    Declare Function B_mnet_m1_set_trigger_comparator Lib "MotionNet.dll" Alias "_mnet_m1_set_trigger_comparator" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal CmpSrc As Integer, ByVal CmpMethod As Integer) As Integer
    Declare Function B_mnet_m1_set_trigger_comparator_data Lib "MotionNet.dll" Alias "_mnet_m1_set_trigger_comparator_data" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal Data As Double) As Integer
    Declare Function B_mnet_m1_get_comparator_data Lib "MotionNet.dll" Alias "_mnet_m1_get_comparator_data" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, POS As Double) As Integer

    Declare Function B_mnet_m1_set_soft_limit Lib "MotionNet.dll" Alias "_mnet_m1_set_soft_limit" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal PLimit As Long, ByVal NLimit As Long) As Integer
    Declare Function B_mnet_m1_disable_soft_limit Lib "MotionNet.dll" Alias "_mnet_m1_disable_soft_limit" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer
    Declare Function B_mnet_m1_enable_soft_limit Lib "MotionNet.dll" Alias "_mnet_m1_enable_soft_limit" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal Action As Integer) As Integer
    Declare Function B_mnet_m1_set_ltc_logic Lib "MotionNet.dll" Alias "_mnet_m1_set_ltc_logic" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal let_logic As Integer) As Integer
    Declare Function B_mnet_m1_steplose_check Lib "MotionNet.dll" Alias "_mnet_m1_steplose_check" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal Tolerance As Integer) As Integer
    Declare Function B_mnet_m1_get_latch_data Lib "MotionNet.dll" Alias "_mnet_m1_get_latch_data" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal LtcNo As Integer, POS As Double) As Integer
    Declare Function B_mnet_m1_start_soft_ltc Lib "MotionNet.dll" Alias "_mnet_m1_start_soft_ltc" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer

    ''Counter Operating
    Declare Function B_mnet_m1_get_command Lib "MotionNet.dll" Alias "_mnet_m1_get_command" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, CMD As Long) As Integer
    Declare Function B_mnet_m1_set_command Lib "MotionNet.dll" Alias "_mnet_m1_set_command" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal CMD As Long) As Integer
    Declare Function B_mnet_m1_reset_command Lib "MotionNet.dll" Alias "_mnet_m1_reset_command" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer
    Declare Function B_mnet_m1_get_position Lib "MotionNet.dll" Alias "_mnet_m1_get_position" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, POS As Long) As Integer
    Declare Function B_mnet_m1_set_position Lib "MotionNet.dll" Alias "_mnet_m1_set_position" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal POS As Long) As Integer
    Declare Function B_mnet_m1_reset_position Lib "MotionNet.dll" Alias "_mnet_m1_reset_position" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer
    Declare Function B_mnet_m1_get_error_counter Lib "MotionNet.dll" Alias "_mnet_m1_get_error_counter" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ErrCnt As Long) As Integer
    Declare Function B_mnet_m1_reset_error_counter Lib "MotionNet.dll" Alias "_mnet_m1_reset_error_counter" (ByVal RingNo As Integer, ByVal SlaveNo As Integer) As Integer

    Declare Function B_mnet_m1_get_current_speed Lib "MotionNet.dll" Alias "_mnet_m1_get_current_speed" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, CurSpeed As Long) As Integer


    '' Homing
    Declare Function B_mnet_m1_set_home_config Lib "MotionNet.dll" Alias "_mnet_m1_set_home_config" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal home_mode As Integer, ByVal org_logic As Integer, ByVal ez_logic As Integer, ByVal ez_count As Integer, ByVal ERC_Out As Integer) As Integer
    Declare Function B_mnet_m1_start_home_move Lib "MotionNet.dll" Alias "_mnet_m1_start_home_move" (ByVal RingNo As Integer, ByVal SlaveNo As Integer, ByVal DIR As Integer) As Integer





End Module
