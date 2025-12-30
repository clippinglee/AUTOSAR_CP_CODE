/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: Tle9210x                                                                                             
*  Content:  Tle9210x family drive
*  Category: Tle92104 Tle92108
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.03.30    clipping            v0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#ifndef _TLE9210X_TYPES_H_
#define _TLE9210X_TYPES_H_
#include "Std_Types.h"
#include "Spi.h"
#include "Dio.h"

#define TLE9210X_BASE_ADDR 0x80u
/***OP bit  Read and write or Read and Clear = 1, Read only = 0***/
#define TLE9210X_OP_RW_OR_R1C 0x01u
#define TLE9210X_OP_READ_ONLY 0x00u

#define TLE9210X_GENCTRL1 0x00u
#define TLE9210X_GENCTRL2 0x01u
#define TLE9210X_VDS1     0x02u
#define TLE9210X_VDS2     0x03u
#define TLE9210X_CCP_BLK1 0x04u
#define TLE9210X_CCP_BLK2_ACT 0x05u
#define TLE9210X_CCP_BLK2_FW  0x05u
#define TLE9210X_HBMODE  0x06u
#define TLE9210X_PWMSET  0x07u
#define TLE9210X_TPRECHG 0x08u
#define TLE9210X_HBIDIAG 0x09u
#define TLE9210X_ST_ICHG 0x0Au
#define TLE9210X_PWM_PCHG_INIT 0x0Au
#define TLE9210X_PWM_ICHG_ACT  0x0Bu
#define TLE9210X_PWM_ICHG_FW   0x0Bu
#define TLE9210X_PWM_IDCHG_ACT 0x0Cu
#define TLE9210X_PWM_PDCHG_INIT 0x0Cu
#define TLE9210X_PWM_ICHGMAX_CCP_BLK3_ACT 0x0Du
#define TLE9210X_PWM_ICHGMAX_CCP_BLK3_FW  0x0Du
#define TLE9210X_TDON_OFF1 0x0Eu
#define TLE9210X_TDON_OFF2 0x0Fu
#define TLE9210X_TDON_OFF3 0x10u

#define TLE9210X_TLE92104 0u
#define TLE9210X_TLE92108 1u

/*****Daisy chain no use lABT = 1******/
#define TLE9210X_DAISY_CHAIN_USER 0u
#define TLE9210X_DAISY_CHAIN_NO_USER 1u

#define TLE9210X_REG_BANK_ON  1u
#define TLE9210X_REG_BANK_OFF 0u

#define TLE9210X_WD_50_MS 0u
#define TLE9210X_WD_200_MS 1u

#define TLE9210X_WD_EN 0u
#define TLE9210X_WD_DIS 1u

#define TLE9210X_GENSTAT 0x11u
#define TLE9210X_DSOV    0x12u
#define TLE9210X_HBVOUT_PWMERR 0x13u
#define TLE9210X_EFF_TDON_OFF1 0x14u
#define TLE9210X_EFF_TDON_OFF2 0x15u
#define TLE9210X_EFF_TDON_OFF3 0x16u
#define TLE9210X_TRISE_FALL1 0x17u
#define TLE9210X_TRISE_FALL2 0x18u
#define TLE9210X_TRISE_FALL3 0x19u
#define TLE9210X_DEVID 0x1Fu

#define TLE9210X_LABT_OFF 0u
#define TLE9210X_LABT_ON 1u

#define TLE9210X_MODE_SLEEP 0u
#define TLE9210X_MODE_NORMAL 1u
#define TLE9210X_MODE_FAIL_SAFE 2u




#define TLE9210X_OUT_STATUS_OFF 0u
#define TLE9210X_OUT_STATUS_LS  1u
#define TLE9210X_OUT_STATUS_HS  2u

#define TLE9210X_HB1 0u
#define TLE9210X_HB2 1u
#define TLE9210X_HB3 2u
#define TLE9210X_HB4 3u
#define TLE9210X_HB5 4u
#define TLE9210X_HB6 5u
#define TLE9210X_HB7 6u
#define TLE9210X_HB8 7u
#define TLE9210X_HB_CHN_MAX 8u

#define TLE9210X_HB_DSM_DH 0u
#define TLE9210X_HB_DSM_CSIN1 1u

#define TLE9210X_HB_DSTH_0_15_V 0u
#define TLE9210X_HB_DSTH_0_20_V 1u
#define TLE9210X_HB_DSTH_0_25_V 2u
#define TLE9210X_HB_DSTH_0_30_V 3u
#define TLE9210X_HB_DSTH_0_40_V 4u
#define TLE9210X_HB_DSTH_0_50_V 5u
#define TLE9210X_HB_DSTH_0_60_V 6u
#define TLE9210X_HB_DSTH_2_00_V 7u

#define TLE9210X_HB_CCPBLK_CCP1_BLANK1 0u
#define TLE9210X_HB_CCPBLK_CCP2_BLANK2 1u
#define TLE9210X_HB_CCPBLK_CCP3_BLANK3 2u
#define TLE9210X_HB_CCPBLK_CCP4_BLANK4 3u

#define TLE9210X_HB_IDIAG_ACTIVATED   1u
#define TLE9210X_HB_IDIAG_DEACTIVATED 0u

#define TLE9210X_HB_ICHGST_CURRENT_1 0u
#define TLE9210X_HB_ICHGST_CURRENT_2 1u

#define TLE9210X_PWM_ENABLE 1u
#define TLE9210X_PWM_DISABLE 0u
#define TLE9210X_PWM_CH1 0u
#define TLE9210X_PWM_CH2 1u
#define TLE9210X_PWM_CH3 2u
#define TLE9210X_PWM_CHN_MAX 3u

#define TLE0210X_PWM_CHARGE_TIME_125NS 0u
#define TLE0210X_PWM_CHARGE_TIME_250NS 1u
#define TLE0210X_PWM_CHARGE_TIME_500NS 2u
#define TLE0210X_PWM_CHARGE_TIME_1000NS 3u


typedef struct
{
    Spi_ChannelType  SpiChannel;
    Spi_SequenceType SpiSequence;
    uint8 u8DaisyChainEn;
    uint8* pu8ChipNum;

}Tle9210x_GroupType;

typedef struct 
{
    uint8 u8ChipId;
    uint8 u8EnPinCtrl;
    boolean REG_BANK;
    boolean WDPER;
    boolean WDDIS;
    uint8 u8CSO1AdcMap;
    uint8 u8CSO2AdcMap;
}Tle9210x_ChipType;

typedef struct 
{
    uint8 u8PwmChn;
    boolean bPwmEn;
    uint8 u8PwmMapChn;
    uint8 u8TurnOnTime;
    uint8 u8TurnOffTime;
    uint8 u8PreChargeTime;
    uint8 u8PreDischargeTime;
}Tle9210x_PwmChnType;

typedef struct 
{
    boolean bHBDrainSrcMonit;
    uint8 u8HBDrainSrcOv;
    uint8 u8HBCCPBLK;
    uint8 u8HBIDIAG;
    uint8 u8HBICHGST;
}Tle9210x_HbChnType;

typedef struct 
{
    uint16 u16GENSTAT;
    uint16 u16DSOV;
    uint16 u16HBVOUT_PWMERR;
    uint16 u16EFF_TDON_OFF1;
    uint16 u16EFF_TDON_OFF2;
    uint16 u16EFF_TDON_OFF3;
    uint16 u16TRISE_FALL1;
    uint16 u16TRISE_FALL2;
    uint16 u16TRISE_FALL3;
    uint16 u16DEVID;
}Tle9210x_GenStsRegType;


#endif
