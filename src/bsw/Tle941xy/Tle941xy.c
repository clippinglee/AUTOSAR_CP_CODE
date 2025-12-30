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
#include "Tle941xy.h"
#include "Tle941xy_Types.h"
#include "Pfm.h"

static uint8 sTle941xy_u8GlobalStatus[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX];
static uint8 sTle941xy_u8PwmDuty[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][TLE941XY_PWM_CHN_MAX];
static uint8 sTle941xy_u8HbOutSts[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][TLE941XY_CHANNEL_MAX];
static PFM_DefectReportState_t sTle941xy_atDiagResult[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX][TLE941XY_CHANNEL_MAX];
static Tle941xy_RegDataType sTle941xy_atRegData[TLE941XY_GROUP_MAX][TLE941XY_CHIP_MAX];
/****************************************************************************************
|     Function Source Code
|***************************************************************************************/

static void Tle941xy_Recovery(uint8 u8GroupId,uint8* pu8RegBuf);
static void Tle941xy_WriteReg(uint8 u8GroupId,uint8* pu8RegBuf, uint8* pu8WtData);
static void Tle941xy_ReadReg( uint8 u8GroupId,uint8* pu8RegBuf,uint8* pu8ReadBuf);

static void Tle941xy_SetHbPwmDutyReg(uint8 u8Group);
static void Tle941xy_ReadDeviceIdReg(uint8 u8Group);
static void Tle941xy_SetFmPwmFreqReg(uint8 u8Group);
static void Tle941xy_SetHbOutputReg(uint8 u8Group);
static void Tle941xy_SetHbModeReg(uint8 u8Group);
static void Tle941xy_ShortDiagnostic(uint8 u8Group);
static void Tle941xy_SetFwOlReg(uint8 u8Group);
static void Tle941xy_OLDiagnostic(uint8 u8Group);
/****************************************************************************************
| NAME:    Tle941xy_WriteReg
| CALLED BY:
| PRECONDITIONS:     NA
| INPUT PARAMETERS:    uint16 buf
| RETURN VALUE:     void
| DESCRIPTION:      recovery tle941xy output (overload)
****************************************************************************************/
static void Tle941xy_WriteReg(uint8 u8GroupId,uint8* pu8RegBuf, uint8* pu8WtData)
{
    uint8 l_u8ChipIndex;
    uint8 l_u8ChipNum;
    uint8 l_au8RcvDataBuf[TLE941XY_CHIP_MAX * 2] = {0};
    uint8 l_au8SndDataBuf[TLE941XY_CHIP_MAX * 2] = {0};

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8GroupId].pu8ChipNum;

    if(l_u8ChipNum > 0u)
    {
        /* Reversal data buffer */
        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            l_au8SndDataBuf[l_u8ChipIndex] =(uint8)(TLE941XY_BASE_ADDR 
                                            | (uint8)(TLE941XY_LABT_OFF << 1u) 
                                            | (uint8)(pu8RegBuf[(l_u8ChipIndex)] << 2u) 
                                            | (uint8)(TLE941XY_WRITE << 7u));
            l_au8SndDataBuf[l_u8ChipIndex + l_u8ChipNum] = pu8WtData[l_u8ChipIndex] ;
        }
        /****The last chip control LABT is 1 whether it is daisy chain communication or not********/
        l_au8SndDataBuf[l_u8ChipNum - 1u] |= (TLE941XY_LABT_ON << 1u);

        /* Send Read  */
        (void)Spi_SetupEB(cTle941xy_atGroupCfg[u8GroupId].SpiChannel, &l_au8SndDataBuf[0], &l_au8RcvDataBuf[0], (l_u8ChipNum * 2));

        (void)Spi_SyncTransmit(cTle941xy_atGroupCfg[u8GroupId].SpiSequence);

        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            sTle941xy_u8GlobalStatus[u8GroupId][l_u8ChipIndex] = l_au8RcvDataBuf[(l_u8ChipNum - l_u8ChipIndex - 1u)];
        }
    }
    else
    {

    }
}

/****************************************************************************************
| NAME:    Tle941xy_ReadReg
| CALLED BY:
| PRECONDITIONS:     NA
| INPUT PARAMETERS:    uint16 buf
| RETURN VALUE:     uint16
| DESCRIPTION:      recovery tle941xy output (overload)
****************************************************************************************/
static void Tle941xy_ReadReg( uint8 u8GroupId,uint8* pu8RegBuf,uint8* pu8ReadBuf)
{
    uint8 l_u8ChipIndex;
    uint8 l_u8ChipNum;
    uint16 l_au8RcvDataBuf[TLE941XY_CHIP_MAX * 2u] = {0};
    uint16 l_au8SndDataBuf[TLE941XY_CHIP_MAX * 2u] = {0};

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8GroupId].pu8ChipNum;

    if(l_u8ChipNum > 0u)
    {
        /* Reversal data buffer */
        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            l_au8SndDataBuf[l_u8ChipIndex] =(uint8)(TLE941XY_BASE_ADDR 
                                            | (uint8)(TLE941XY_LABT_OFF << 1u) 
                                            | (uint8)(pu8RegBuf[(l_u8ChipIndex)] << 2u) 
                                            | (uint8)(TLE941XY_READ << 7u));
        }
        /****The last chip control LABT is 1 whether it is daisy chain communication or not********/
        l_au8SndDataBuf[l_u8ChipNum - 1u] |= (uint8)(TLE941XY_LABT_ON << 1u);

        /* Send Read  */
        (void)Spi_SetupEB(cTle941xy_atGroupCfg[u8GroupId].SpiChannel, &l_au8SndDataBuf[0], &l_au8RcvDataBuf[0], (l_u8ChipNum * 2u));

        (void)Spi_SyncTransmit(cTle941xy_atGroupCfg[u8GroupId].SpiSequence);

        for(l_u8ChipIndex = 0u; l_u8ChipIndex < l_u8ChipNum; l_u8ChipIndex++)
        {
            pu8ReadBuf[l_u8ChipIndex] = l_au8RcvDataBuf[(l_u8ChipNum * 2u - l_u8ChipIndex - 1u)];
        }
    }
    else
    {

    }

}

static void Tle941xy_SetHbOutputReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;
    /***OUT1-OUT4**/
    for(j = 0u;j<l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_HB_ACT_1_CTRL;
        l_au8DataBuf[j] = (sTle941xy_u8HbOutSts[u8Group][j][0] 
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][1] << 2u)
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][2] << 4u)
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][3] << 6u));
    }
    sTle941xy_atRegData[u8Group][j].HB_ACT_1_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#if((TLE941XY_TLE94106_CHIP_EN == STD_ON)||(TLE941XY_TLE94108_CHIP_EN == STD_ON)||(TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT5-OUT8**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_HB_ACT_2_CTRL;
        l_au8DataBuf[j] = (sTle941xy_u8HbOutSts[u8Group][j][4] 
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][5] << 2u)
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][6] << 4u)
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][7] << 6u));
    }
    sTle941xy_atRegData[u8Group][j].HB_ACT_2_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#endif
#if((TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT9-OUT12**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_HB_ACT_3_CTRL;
        l_au8DataBuf[j] = (sTle941xy_u8HbOutSts[u8Group][j][8] 
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][9] << 2u)
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][10] << 4u)
                        | (uint8)(sTle941xy_u8HbOutSts[u8Group][j][11] << 6u));
    }
    sTle941xy_atRegData[u8Group][j].HB_ACT_3_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#endif
}

static void Tle941xy_SetHbModeReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;
    /***OUT1-OUT4**/
    for(j = 0u;j<l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_HB_MODE_1_CTRL;
        l_au8DataBuf[j] = (cTle941xy_au8ChnModeCfg[u8Group][j][0] 
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][1] << 2u)
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][2] << 4u)
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][3] << 6u));
    }
    sTle941xy_atRegData[u8Group][j].HB_MODE_1_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#if((TLE941XY_TLE94106_CHIP_EN == STD_ON)||(TLE941XY_TLE94108_CHIP_EN == STD_ON)||(TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT5-OUT8**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_HB_MODE_2_CTRL;
        l_au8DataBuf[j] = (cTle941xy_au8ChnModeCfg[u8Group][j][4] 
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][5] << 2u)
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][6] << 4u)
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][7] << 6u));
    }
    sTle941xy_atRegData[u8Group][j].HB_MODE_2_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#endif
#if((TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT9-OUT12**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_HB_MODE_3_CTRL;
        l_au8DataBuf[j] = (cTle941xy_au8ChnModeCfg[u8Group][j][8] 
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][9] << 2u)
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][10] << 4u)
                        | (uint8)(cTle941xy_au8ChnModeCfg[u8Group][j][11] << 6u));
    }
    sTle941xy_atRegData[u8Group][j].HB_MODE_3_CTRL = l_au8DataBuf[j];   
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#endif
}

static void Tle941xy_SetFmPwmFreqReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;
    /***OUT1-OUT4**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        if((cTle941xy_atChipCfg[u8Group][j].Tle941xyChipId == TLE941XY_TLE94103) ||(cTle941xy_atChipCfg[u8Group][j].Tle941xyChipId == TLE941XY_TLE94104))
        {
            l_au8RegBuf[j] = TLE941XY_FM_CLK_CTRL;
        }
        else
        {
            l_au8RegBuf[j] = TLE941XY_PWM_CH_FREQ_CTRL;
        }
        l_au8DataBuf[j] = (cTle941xy_atChipFmPwmFreqCfg[u8Group][j].u8FmFreq
                        | (uint8)(cTle941xy_atChipFmPwmFreqCfg[u8Group][j].u8Pwm1Freq << 2u)
                        | (uint8)(cTle941xy_atChipFmPwmFreqCfg[u8Group][j].u8Pwm2Freq << 4u)
                        | (uint8)(cTle941xy_atChipFmPwmFreqCfg[u8Group][j].u8Pwm3Freq << 6u));
    }
    sTle941xy_atRegData[u8Group][j].PWM_CH_FREQ_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
}

static void Tle941xy_ReadDeviceIdReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;
    for(j = 0u;j<l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_CONFIG_CTRL;
    }
    sTle941xy_atRegData[u8Group][j].CONFIG_CTRL = l_au8DataBuf[j];
    Tle941xy_ReadReg(u8Group,l_au8RegBuf,l_au8DataBuf);
}


static void Tle941xy_SetHbPwmDutyReg(uint8 u8Group)
{
#if((TLE941XY_TLE94106_CHIP_EN == STD_ON)||(TLE941XY_TLE94108_CHIP_EN == STD_ON)||(TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    uint8 j;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;

    for(j = 0u;j<l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_PWM1_DC_CTRL;
        l_au8DataBuf[j] = sTle941xy_u8PwmDuty[u8Group][j][0];
    }
    sTle941xy_atRegData[u8Group][j].PWM1_DC_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);

    for(j = 0u;j<l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_PWM2_DC_CTRL;
        l_au8DataBuf[j] = sTle941xy_u8PwmDuty[u8Group][j][1];
    }
    sTle941xy_atRegData[u8Group][j].PWM2_DC_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);

    for(j = 0u;j<l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_PWM3_DC_CTRL;
        l_au8DataBuf[j] = sTle941xy_u8PwmDuty[u8Group][j][2];
    }
    sTle941xy_atRegData[u8Group][j].PWM3_DC_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);

#endif
}

static void Tle941xy_SetFwOlReg(uint8 u8Group)
{
    uint8 j;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;

#if((TLE941XY_TLE94106_CHIP_EN == STD_ON)||(TLE941XY_TLE94108_CHIP_EN == STD_ON)||(TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT5-OUT8**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_FW_OL_CTRL;
        l_au8DataBuf[j] = (cTle941xy_abChipHS1And2LedModeCfg[u8Group][j][0] 
                        | (uint8)(cTle941xy_abChipHS1And2LedModeCfg[u8Group][j][1] << 1u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][0] << 2u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][1] << 3u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][2] << 4u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][3] << 5u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][4] << 6u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][5] << 7u));
    }
    sTle941xy_atRegData[u8Group][j].FW_OL_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#endif
#if((TLE941XY_TLE94103_CHIP_EN == STD_ON)||(TLE941XY_TLE94104_CHIP_EN == STD_ON)||(TLE941XY_TLE94108_CHIP_EN == STD_ON)||(TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT9-OUT12**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_FW_CTRL;
        l_au8DataBuf[j] = (cTle941xy_abChipFreeWheelingCfg[u8Group][j][6] 
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][7] << 1u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][8] << 2u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][9] << 3u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][10] << 4u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][11] << 5u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][0] << 6u)
                        | (uint8)(cTle941xy_abChipFreeWheelingCfg[u8Group][j][0] << 7u));
    }
    sTle941xy_atRegData[u8Group][j].FW_CTRL = l_au8DataBuf[j];
    Tle941xy_WriteReg(u8Group,l_au8RegBuf,l_au8DataBuf);
#endif
}

/****************************************************************************************
| NAME:    Tle941xy_ShortDiagnostic
| CALLED BY:
| PRECONDITIONS:     NA
| INPUT PARAMETERS:    void
| RETURN VALUE:     void
| DESCRIPTION:      recovery Tle941xy output (overload)
****************************************************************************************/
static void Tle941xy_ShortDiagnostic(uint8 u8Group)
{
    uint8 j;
    uint8 k;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;
    uint8 l_u8DisplacementLen;
    uint8 l_u8ChipShortFlag;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;
    /***OUT1-OUT4**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_SYS_DIAG_2;
    }
    Tle941xy_ReadReg(u8Group,l_au8RegBuf,l_au8DataBuf);
    sTle941xy_atRegData[u8Group][j].SYS_DIAG_2 = l_au8DataBuf[j];
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        for(j = 0u;k < 4u;k++)
        {
            l_u8DisplacementLen = (uint8)(k * 2u);
            if((l_au8RegBuf[j] && (0x03u << l_u8DisplacementLen)) != 0x00u)
            {
                l_u8ChipShortFlag++;
                if(sTle941xy_u8HbOutSts[u8Group][j][k] == TLE941XY_OUT_STATUS_LS)
                {
                    sTle941xy_atDiagResult[u8Group][j][k].eShort2Vcc = PFM_DDS_POS;
                }
                else
                {
                    sTle941xy_atDiagResult[u8Group][j][k].eShort2Gnd = PFM_DDS_POS;
                }
            }
            else
            {
                sTle941xy_atDiagResult[u8Group][j][k].eShort2Vcc = PFM_DDS_NEG;
                sTle941xy_atDiagResult[u8Group][j][k].eShort2Gnd = PFM_DDS_NEG;
            }
        }
    }
    if(l_u8ChipShortFlag > 0u)
    {
        Tle941xy_Recovery(u8Group,l_au8RegBuf);
    }
    else
    {
        
    }

#if((TLE941XY_TLE94106_CHIP_EN == STD_ON)||(TLE941XY_TLE94108_CHIP_EN == STD_ON)||(TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    l_u8ChipShortFlag = 0u;
    /***OUT5-OUT8**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_SYS_DIAG_3;
    }
    Tle941xy_ReadReg(u8Group,l_au8RegBuf,l_au8DataBuf);
    sTle941xy_atRegData[u8Group][j].SYS_DIAG_3 = l_au8DataBuf[j];
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        for(j = 0u;k < 4u;k++)
        {
            l_u8DisplacementLen = (uint8)(k * 2u);
            if((l_au8RegBuf[j] && (0x03u << l_u8DisplacementLen)) != 0x00u)
            {
                l_u8ChipShortFlag++;
                if(sTle941xy_u8HbOutSts[u8Group][j][k + 4u] == TLE941XY_OUT_STATUS_LS)
                {
                    sTle941xy_atDiagResult[u8Group][j][k + 4u].eShort2Vcc = PFM_DDS_POS;
                }
                else
                {
                    sTle941xy_atDiagResult[u8Group][j][k + 4u].eShort2Gnd = PFM_DDS_POS;
                }
            }
            else
            {
                sTle941xy_atDiagResult[u8Group][j][k + 4u].eShort2Vcc = PFM_DDS_NEG;
                sTle941xy_atDiagResult[u8Group][j][k + 4u].eShort2Gnd = PFM_DDS_NEG;
            }
        }
    }
    if(l_u8ChipShortFlag > 0u)
    {
        Tle941xy_Recovery(u8Group,l_au8RegBuf);
    }
    else
    {
        
    }
#endif
#if((TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    l_u8ChipShortFlag = 0u;
    /***OUT9-OUT12**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_SYS_DIAG_4;
    }
    Tle941xy_ReadReg(u8Group,l_au8RegBuf,l_au8DataBuf);
    sTle941xy_atRegData[u8Group][j].SYS_DIAG_4 = l_au8DataBuf[j];
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        for(j = 0u;k < 4u;k++)
        {
            l_u8DisplacementLen = (uint8)(k * 2u);
            if((l_au8RegBuf[j] && (0x03u << l_u8DisplacementLen)) != 0x00u)
            {
                l_u8ChipShortFlag++;
                if(sTle941xy_u8HbOutSts[u8Group][j][k + 8u] == TLE941XY_OUT_STATUS_LS)
                {
                    sTle941xy_atDiagResult[u8Group][j][k + 8u].eShort2Vcc = PFM_DDS_POS;
                }
                else
                {
                    sTle941xy_atDiagResult[u8Group][j][k + 8u].eShort2Gnd = PFM_DDS_POS;
                }
            }
            else
            {
                sTle941xy_atDiagResult[u8Group][j][k + 8u].eShort2Vcc = PFM_DDS_NEG;
                sTle941xy_atDiagResult[u8Group][j][k + 8u].eShort2Gnd = PFM_DDS_NEG;
            }
        }
        if(l_u8ChipShortFlag > 0u)
        {
            Tle941xy_Recovery(u8Group,l_au8RegBuf);
        }
        else
        {
            
        }
    }
#endif
}

static void Tle941xy_Recovery(uint8 u8GroupId,uint8* pu8RegBuf)
{
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};

    (void)memset(l_au8DataBuf,1u,sizeof(l_au8DataBuf));
    /****write 1 clear 0******/
    Tle941xy_WriteReg(u8GroupId,pu8RegBuf,l_au8DataBuf);
}

/****************************************************************************************
| NAME:    Tle941xy_OLDiagnostic
| CALLED BY:
| PRECONDITIONS:     NA
| INPUT PARAMETERS:    void
| RETURN VALUE:     void
| DESCRIPTION:      
****************************************************************************************/
static void Tle941xy_OLDiagnostic(uint8 u8Group)
{

    uint8 j;
    uint8 k;
    uint8 l_au8RegBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_au8DataBuf[TLE941XY_CHIP_MAX] = {0};
    uint8 l_u8ChipNum;
    uint8 l_u8DisplacementLen;

    l_u8ChipNum = *cTle941xy_atGroupCfg[u8Group].pu8ChipNum;
    /***OUT1-OUT4**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_SYS_DIAG_5;
    }
    Tle941xy_ReadReg(u8Group,l_au8RegBuf,l_au8DataBuf);
    sTle941xy_atRegData[u8Group][j].SYS_DIAG_5 = l_au8DataBuf[j];
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        for(j = 0u;k < 4u;k++)
        {
            l_u8DisplacementLen = (uint8)(k * 2u);
            if((l_au8RegBuf[j] && (0x03u << l_u8DisplacementLen)) != 0x00u)
            { 
                sTle941xy_atDiagResult[u8Group][j][k].eOpenLoad = PFM_DDS_POS; 
            }
            else
            {
                sTle941xy_atDiagResult[u8Group][j][k].eOpenLoad = PFM_DDS_NEG;
            }
        }
    }

#if((TLE941XY_TLE94106_CHIP_EN == STD_ON)||(TLE941XY_TLE94108_CHIP_EN == STD_ON)||(TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT5-OUT8**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_SYS_DIAG_6;
    }
    Tle941xy_ReadReg(u8Group,l_au8RegBuf,l_au8DataBuf);
    sTle941xy_atRegData[u8Group][j].SYS_DIAG_6 = l_au8DataBuf[j];
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        for(j = 0u;k < 4u;k++)
        {
            l_u8DisplacementLen = (uint8)(k * 2u);
            if((l_au8RegBuf[j] && (0x03u << l_u8DisplacementLen)) != 0x00u)
            { 
                sTle941xy_atDiagResult[u8Group][j][k + 4u].eOpenLoad = PFM_DDS_POS; 
            }
            else
            {
                sTle941xy_atDiagResult[u8Group][j][k + 4u].eOpenLoad = PFM_DDS_NEG;
            }
        }
    }
#endif
#if((TLE941XY_TLE94110_CHIP_EN == STD_ON)||(TLE941XY_TLE94112_CHIP_EN == STD_ON))
    /***OUT9-OUT12**/
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        l_au8RegBuf[j] = TLE941XY_SYS_DIAG_7;
    }
    Tle941xy_ReadReg(u8Group,l_au8RegBuf,l_au8DataBuf);
    sTle941xy_atRegData[u8Group][j].SYS_DIAG_7 = l_au8DataBuf[j];
    for(j = 0u;j < l_u8ChipNum;j++)
    {
        for(j = 0u;k < 4u;k++)
        {
            l_u8DisplacementLen = (uint8)(k * 2u);
            if((l_au8RegBuf[j] && (0x03u << l_u8DisplacementLen)) != 0x00u)
            { 
                sTle941xy_atDiagResult[u8Group][j][k + 8u].eOpenLoad = PFM_DDS_POS; 
            }
            else
            {
                sTle941xy_atDiagResult[u8Group][j][k + 8u].eOpenLoad = PFM_DDS_NEG;
            }
        }
    }
#endif
}

void Tle941xy_Init(void)
{
    uint8 i;
    uint8 j;
    uint8 l_u8ChipNum;

    (void)memset(sTle941xy_u8HbOutSts,0u,sizeof(sTle941xy_u8HbOutSts));
    for(i = 0u;i < TLE941XY_GROUP_MAX;i++)
    {
        l_u8ChipNum = *cTle941xy_atGroupCfg[i].pu8ChipNum;
        for(j = 0u;j < l_u8ChipNum;j++)
        {
            Dio_WriteChannel(cTle941xy_atChipCfg[i][j].u8ChipEnPin, STD_ON);
        }
        Tle941xy_ReadDeviceIdReg(i);
        Tle941xy_SetHbModeReg(i);
        Tle941xy_SetFwOlReg(i);
        Tle941xy_SetFmPwmFreqReg(i);
        Tle941xy_SetHbOutputReg(i);
    }
}

void Tle941xy_MainFunction(void)
{
    uint8 i;
    for(i = 0u;i < TLE941XY_GROUP_MAX;i++)
    {
        Tle941xy_ShortDiagnostic(i);
        Tle941xy_OLDiagnostic(i);
        Tle941xy_SetHbPwmDutyReg(i);
        Tle941xy_SetHbOutputReg(i);
    }
}

void Tle941xy_DeInit(void)
{
    uint8 i;
    (void)memset(sTle941xy_u8HbOutSts,0u,sizeof(sTle941xy_u8HbOutSts));
    for(i = 0u;i < TLE941XY_GROUP_MAX;i++)
    {
        Tle941xy_SetHbOutputReg(i);
    }
}

/****************************************************************************************
| NAME:    Tle941xy_WriteOhbChn
| CALLED BY:     output layer
| PRECONDITIONS:     NA
| INPUT PARAMETERS:    uint8 channel, uint16 Value
| RETURN VALUE:     void
| DESCRIPTION:      The interface to write value to a DO channel.
 ****************************************************************************************/
void Tle941xy_WriteHbChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8ChnId, uint8 u8Val)
{
    if((u8GroupId < (uint8)TLE941XY_GROUP_MAX)
    &&(u8ChipId < (uint8)TLE941XY_CHIP_MAX)
    &&(u8ChnId < (uint8)TLE941XY_CHANNEL_MAX))
    {
        sTle941xy_u8HbOutSts[u8GroupId][u8ChipId][u8ChnId] = u8Val;
    }
}

void Tle941xy_WritePwmChn(uint8 u8GroupId, uint8 u8ChipId,uint8 u8PwmChn, uint8 u8Val)
{
    if((u8GroupId < (uint8)TLE941XY_GROUP_MAX)
    &&(u8ChipId < (uint8)TLE941XY_CHIP_MAX)
    &&(u8PwmChn < (uint8)TLE941XY_PWM_CHN_MAX))
    {
        sTle941xy_u8PwmDuty[u8GroupId][u8ChipId][u8PwmChn] = u8Val;
    }
}
