/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: Tle941xy                                                                                             
*  Content:  Tle941xy family drive
*  Category: Tle94103 Tle94104 Tle94106 Tle94108 Tle94110 Tle94112 
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.03.30    clipping            v0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#ifndef _TLE941XY_TYPES_H_
#define _TLE941XY_TYPES_H_

#include "Std_Types.h"
#include "Spi.h"
#include "Dio.h"

#define TLE941XY_TLE94103 5u
#define TLE941XY_TLE94104 4u
#define TLE941XY_TLE94106 3u
#define TLE941XY_TLE94108 2u
#define TLE941XY_TLE94110 1u
#define TLE941XY_TLE94112 0u


#define TLE941XY_CHANNEL_01 0u
#define TLE941XY_CHANNEL_02 1u
#define TLE941XY_CHANNEL_03 2u
#define TLE941XY_CHANNEL_04 3u
#define TLE941XY_CHANNEL_05 4u
#define TLE941XY_CHANNEL_06 5u
#define TLE941XY_CHANNEL_07 6u
#define TLE941XY_CHANNEL_08 7u
#define TLE941XY_CHANNEL_09 8u
#define TLE941XY_CHANNEL_10 9u
#define TLE941XY_CHANNEL_11 10u
#define TLE941XY_CHANNEL_12 11u
#define TLE941XY_CHANNEL_MAX 12u

#define TLE941XY_TLE94103_CHN_MAX 3u
#define TLE941XY_TLE94104_CHN_MAX 4u
#define TLE941XY_TLE94106_CHN_MAX 6u
#define TLE941XY_TLE94108_CHN_MAX 8u
#define TLE941XY_TLE94110_CHN_MAX 10u
#define TLE941XY_TLE94112_CHN_MAX TLE941XY_CHANNEL_MAX



/*******TLE941XY Register*******/
#define TLE941XY_BASE_ADDR 0x01u

/*****Daisy chain no use lABT = 1******/
#define TLE941XY_DAISY_CHAIN_USER 0u
#define TLE941XY_DAISY_CHAIN_NO_USER 1u

#define TLE941XY_LABT_OFF 0u
#define TLE941XY_LABT_ON 1u

#define TLE941XY_READ 0u
#define TLE941XY_WRITE 1u

#define TLE941XY_HB_ACT_1_CTRL 0x00u
#define TLE941XY_HB_ACT_2_CTRL 0x10u
#define TLE941XY_HB_ACT_3_CTRL 0x08u
#define TLE941XY_HB_MODE_1_CTRL 0x18u
#define TLE941XY_HB_MODE_2_CTRL 0x04u
#define TLE941XY_HB_MODE_3_CTRL 0x14u
#define TLE941XY_PWM_CH_FREQ_CTRL 0x0Cu
#define TLE941XY_PWM1_DC_CTRL 0x1Cu
#define TLE941XY_PWM2_DC_CTRL 0x02u
#define TLE941XY_PWM3_DC_CTRL 0x12u
#define TLE941XY_FW_OL_CTRL 0x0Au
#define TLE941XY_FW_CTRL 0x1Au
#define TLE941XY_CONFIG_CTRL 0x19u
#define TLE941XY_FM_CLK_CTRL 0x0Cu
#define TLE941XY_OLBLK_CTRL 0x1Au

#define TLE941XY_SYS_DIAG_1 0x06u
#define TLE941XY_SYS_DIAG_2 0x16u
#define TLE941XY_SYS_DIAG_3 0x0Eu
#define TLE941XY_SYS_DIAG_4 0x1Eu
#define TLE941XY_SYS_DIAG_5 0x01u
#define TLE941XY_SYS_DIAG_6 0x11u
#define TLE941XY_SYS_DIAG_7 0x09u

#define TLE941XY_CHN_DUMMY     0u
#define TLE941XY_CHN_CTRL_DO   0U
#define TLE941XY_CHN_CTRL_PWM1 1U
#define TLE941XY_CHN_CTRL_PWM2 2U
#define TLE941XY_CHN_CTRL_PWM3 3U

#define TLE941XY_PWM_CHN_MAX 3U

#define TLE941XY_PWM_OFF         0u
#define TLE941XY_PWM_FREQ_080HZ  1u
#define TLE941XY_PWM_FREQ_100HZ  2u
#define TLE941XY_PWM_FREQ_200HZ  3u

#define TLE941XY_FM_MOD_OFF        0u   /*OFF*/
#define TLE941XY_FM_MOD_FREQ_15KHZ 1u   /*15.625kHz*/
#define TLE941XY_FM_MOD_FREQ_31KHZ 2u   /*31.25kHz*/
#define TLE941XY_FM_MOD_FREQ_62KHZ 3u   /*62.5kHz*/


#define TLE941XY_OUT_STATUS_OFF 0u
#define TLE941XY_OUT_STATUS_LS  1u
#define TLE941XY_OUT_STATUS_HS  2u

#define TLE941XY_CHN_FW_ON  1u
#define TLE941XY_CHN_FW_OFF 0u

#define TLE941XY_CHN_LED_ON  1u
#define TLE941XY_CHN_LED_OFF  0u


typedef struct 
{
    uint8 u8FmFreq;
    uint8 u8Pwm1Freq;
    uint8 u8Pwm2Freq;
    uint8 u8Pwm3Freq;
}Tle941xy_PwmType;

typedef struct
{
    Spi_ChannelType  SpiChannel;
    Spi_SequenceType SpiSequence;
    uint8 u8DaisyChainEn;
    uint8* pu8ChipNum;
}Tle941xy_GroupType;

typedef struct
{
    uint8 au8InputType[TLE941XY_CHANNEL_MAX];
}Tle941xy_ChnType;

typedef struct
{
    uint8 Tle941xyChipId;
    Dio_ChannelType u8ChipEnPin;
}Tle941xy_ChipType;

typedef struct
{
    uint8 HB_ACT_1_CTRL;
    uint8 HB_ACT_2_CTRL;
    uint8 HB_ACT_3_CTRL;
    uint8 HB_MODE_1_CTRL;
    uint8 HB_MODE_2_CTRL;
    uint8 HB_MODE_3_CTRL;
    uint8 PWM_CH_FREQ_CTRL;
    uint8 PWM1_DC_CTRL;
    uint8 PWM2_DC_CTRL;
    uint8 PWM3_DC_CTRL;
    uint8 FW_OL_CTRL;
    uint8 FW_CTRL;
    uint8 CONFIG_CTRL;
    uint8 SYS_DIAG_1;
    uint8 SYS_DIAG_2;
    uint8 SYS_DIAG_3;
    uint8 SYS_DIAG_4;
    uint8 SYS_DIAG_5;
    uint8 SYS_DIAG_6;
    uint8 SYS_DIAG_7;
}Tle941xy_RegDataType;

#endif



