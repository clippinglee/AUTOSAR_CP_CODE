/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: Vn7x                                                                                             
*  Content:  Vn7 family drive
*  Category:                                                                              				          
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.03.31    clipping            V0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#include "Vn7x.h"
#include "Vn7x_Types.h"
#include "Vn7x_HwCfg.h"

#include "AdcIf.h"
#include "PFM.h"
#include <string.h>
#include "LiBool.h"
/* PRQA S 0314 EOF*/

/*******************************************************************************
**  Local  variable definitions
*******************************************************************************/

/* the fault status of all channels*/
static PFM_DefectReportState_t sVn7x_atDiagResult[VN7X_ID_MAX];

/* record of current diagnostic channel selection */
static uint8 sVn7x_u8ChnSel = VN7X_DAIG_SEL_CHN_ZERO;

/*******************************************************************************
**  Global  variable definitions
*******************************************************************************/
static uint16 sVn7x_au16PwmOutDuty[VN7X_ID_MAX];
static boolean sVn7x_abDoValue[VN7X_ID_MAX];
/* the record of ON/OFF status of all channels */
static uint32 sVn7x_u32ChnSts;
/* ADC value of feedback diagnostic signals of all channels */
static uint16 gVn7x_au16DiagAdcV[VN7X_ID_MAX];

#define VN7X_GETCHANSTATE(port)    GETBIT_U32(sVn7x_u32ChnSts, port)
/*******************************************************************************
**  Local  Function definitions
*******************************************************************************/
static void Vn7x_DiagChanSw(void);
static void Vn7x_GetDiagAdVal(void);
static void Vn7x_DiagHandle(void);
static boolean Vn7x_JudgePwmDuty(uint16 u16NominalValue);
static void Vn7x_WriteOutput(void);
/*******************************************************************************
**  Global  Function definitions
*******************************************************************************/

/****************************************************************
 process: Vn7x_vDiagHandle
 purpose: The diagnostic handler for VN7X module. Invoked every 10ms.
          Figure out if openload, short to battery or short to ground
          fault has occured. Always report current fault status to 
          PFM module. Debounce is performed in PFM module.
 ****************************************************************/
static void Vn7x_DiagHandle(void)
{
    boolean l_bChanState;
    uint8   l_u8Port;
    uint8   l_u8DiagMode;
    uint8   l_u8DiagChn;
    uint16  l_u16DiagRaw;
    PFM_PhysicalId_e l_eFid; 
    
    /* Go through all channels and perform diagnostic operation.
       Report diagnosing result to PFM. */
    for( l_u8Port = 0u; l_u8Port < (uint8)VN7X_ID_MAX; l_u8Port++ )
    {
        /* Get the index of current channel in PFM module*/
        l_bChanState = (boolean)(VN7X_GETCHANSTATE(l_u8Port) ? TRUE : FALSE);
        /* if diagnosing channel selection equals this channel (channel 0 or channel 1),
           which means the ADC sample value belongs to this channel, diagnosing can be 
           performed. Otherwise the diagnosing of this channel should wait until next cycle*/
        if( (l_u8DiagChn == sVn7x_u8ChnSel) 
            && (GETBIT_U32(sVn7x_u32ChnSts,l_u8Port)
            || (l_bChanState == (boolean)FALSE)))
        {
            if(gVn7x_au16DiagAdcV[l_u8Port] <= cVn7x_atChannelInputCfg[l_u8Port].u16OLDiagAdcVal)
            {
                sVn7x_atDiagResult[l_u8Port].eOpenLoad  = PFM_DDS_POS;
            }
            else if(gVn7x_au16DiagAdcV[l_u8Port] >= cVn7x_atChannelInputCfg[l_u8Port].u16ShortDiagAdcVal)
            {
                sVn7x_atDiagResult[l_u8Port].eShort2Gnd  = PFM_DDS_POS;
            }
            else
            {
                sVn7x_atDiagResult[l_u8Port].eOpenLoad  = PFM_DDS_ING;
                sVn7x_atDiagResult[l_u8Port].eShort2Vcc = PFM_DDS_ING;
            }
            sVn7x_atDiagResult[l_u8Port].eShort2Gnd = PFM_DDS_ING;
        }
        else   /* If this channel is not selected as feedback source, wait for next cycle */
        {
            sVn7x_atDiagResult[l_u8Port].eOpenLoad  = PFM_DDS_ING;
            sVn7x_atDiagResult[l_u8Port].eShort2Vcc = PFM_DDS_ING;
            sVn7x_atDiagResult[l_u8Port].eShort2Gnd = PFM_DDS_ING;
        }
        Pfm_DefectReport(l_eFid, sVn7x_atDiagResult[l_u8Port].eOpenLoad, sVn7x_atDiagResult[l_u8Port].eShort2Vcc, sVn7x_atDiagResult[l_u8Port].eShort2Gnd);
    }
}

/**************************************************************************
 process: Vn7x_vChanSw
 purpose: VN7X has 2 output channels but only 1 feedback diagnostic channel.
          So the diagnostic feedback channel should switch between 2 outputs.
          This function is called every 10 ms and switch feedback channel to 
          the other output.
 **************************************************************************/
static void Vn7x_DiagChanSw(void)
{
    /* select diagnostic channel here. */
    /* all diag channel select pins of all HSD chips are connected together*/
    VN7X_DISABLE_PEM_TRIGGER_ADC
    if(sVn7x_u8ChnSel < VN7X_DAIG_SEL_CHN_MAX)
    {
        sVn7x_u8ChnSel++;
    }
    else
    {
        sVn7x_u8ChnSel = VN7X_DAIG_SEL_CHN_ZERO;
    }

    VN7X_ENABLE_PWM_TRIGGER_ADC;

}

/****************************************************************
 process: Vn7x_vGetDiagAdVal
 purpose: This function records diagnostic raw values and external 
          control signal from ADC buffer.
          Some lamps have redundant control which is realized by a 
          hardwire. If VN7X channel is off while external control 
          is on, there is feedback current on diagnostic pin anyway.
          This function recognises this situation by reading ADC
          value of corresponding hardwire control.
 ****************************************************************/
static void Vn7x_GetDiagAdVal(void)
{
    uint8 l_u8Port;
    uint8 l_u8DiagChan;

    for (l_u8Port = 0u; l_u8Port < (uint8)VN7X_ID_MAX; l_u8Port ++)
    {
        /* If selected channel is being diagnosed, the ADC value of feedback diagnostic channel and
            external hardwired control signal detection(if not applicable, set to 0) should be retrieved*/
        if (l_u8DiagChan == sVn7x_u8ChnSel)
        {
            /* Get the AD value of diagnostic feedback */
            gVn7x_au16DiagAdcV[l_u8Port] = AdcIf_GetAdcValue(cVn7x_au8AdcEidMap[l_u8Port]);
        }
        else
        {
        }
    }
}

/****************************************************************
 process: Vn7x_vTurnOffAll
 purpose: Turn off all output channel ae weel as re-initilizing
          the status record "gVn7x_u32ChanStat".
 ****************************************************************/
void Vn7x_TurnOffAll(void)
{
    uint8 l_u8Port;
    for(l_u8Port = 0u; l_u8Port < (uint8)VN7X_ID_MAX; l_u8Port++)
    {
        Vn7x_WriteDoChn(l_u8Port, 0u);
    }
    /* update all channel state record to off */
    sVn7x_u32ChnSts = STD_OFF;   
    Vn7x_WriteOutput();
}

static void Vn7x_WriteOutput(void)
{
    uint8 i;
    for(i = 0;i < (uint8)VN7X_ID_MAX;i++)
    {
        if(VN7X_PWM == cVn7x_atChannelInputCfg[i].eVn7x_Type)
        {
            Pwm_SetDutyCycle(cVn7x_atChannelInputCfg[i].u8Vn7xPwmCntrl, sVn7x_au16PwmOutDuty[i]);
        }
        else if( VN7X_DIO == cVn7x_atChannelInputCfg[i].eVn7x_Type)
        {
            Dio_WriteChannel(cVn7x_atChannelInputCfg[i].u8Vn7xDioInput, sVn7x_abDoValue[i]);
        }
        else
        {
            /*do nothing*/
        }
    }

}

/****************************************************************
 process: Vn7x_v10ms
 purpose: VN7X module main function. invuked every 10 ms.
 ****************************************************************/
void Vn7x_MainFunction(void)
{
    Vn7x_GetDiagAdVal();
    Vn7x_DiagHandle();
    Vn7x_WriteOutput();
    Vn7x_DiagChanSw();
}

/****************************************************************
 process: Vn7x_vInit
 purpose: Init process.
 ****************************************************************/
void Vn7x_Init(void)
{
    uint8 i;
    /* initialize the global diagnostic variables */
    (void)memset((void *)gVn7x_au16DiagAdcV, 0, sizeof(uint16) * (uint8)VN7X_ID_MAX);                /* initialize AD values of all diagnostic channels */
    (void)memset((void *)sVn7x_atDiagResult,0,sizeof(PFM_DefectReportState_t) * (uint8)VN7X_ID_MAX); /* all diagnostic result = PFM_DDS_ING(0) */
    /* Select output 0 as feedback source at first place */
    for(i = 0u;i < VN7X_ID_MAX;i++)
    {
        Dio_WriteChannel(cVn7x_atChannelInputCfg[i].u8Vn7xDioSEn, sVn7x_abDoValue[i]);
    }

    sVn7x_u8ChnSel = VN7X_DAIG_SEL_CHN_ZERO;
    /* turn off all outputs and initialize state record to all-off */
    Vn7x_TurnOffAll();  
}
/****************************************************************
 process: Vn7x_vDeInit
 purpose: Put all VN7X chips to initialized state as well as 
          turning off all outputs.
          If application layer does not update VN7X on/off status right 
          after waking up, the outputs should not be closed here.
 ****************************************************************/
void Vn7x_DeInit(void)
{  
    /* turn off all outputs and initialize state record to all-off */
    Vn7x_TurnOffAll();
    /* initialize the global diagnostic variables */
    (void)memset((void *)gVn7x_au16DiagAdcV, 0, sizeof(uint16) * (uint8)VN7X_ID_MAX);                  /* zero all diagnostic signal AD value */

    (void)memset((void *)sVn7x_atDiagResult, 0, sizeof(PFM_DefectReportState_t)*  (uint8)VN7X_ID_MAX); /* all diagnostic result = PFM_DDS_ING(0) */
    /* Switch feedback source to channel 0 */
    sVn7x_u8ChnSel = VN7X_DAIG_SEL_CHN_ZERO;
}

/****************************************************************
 process: Vn7x_vWriteDoChn
 purpose: Write Channel Value.
 ****************************************************************/
void Vn7x_WriteDoChn(uint8 u8Chn, uint16 u16Val)
{
    if(VN7X_PWM == cVn7x_atChannelInputCfg[u8Chn].eVn7x_Type)
    {
        sVn7x_au16PwmOutDuty[u8Chn] = u16Val;
    }
    else if( VN7X_DIO== cVn7x_atChannelInputCfg[u8Chn].eVn7x_Type)
    {
        sVn7x_abDoValue[u8Chn] = (boolean)u16Val;
    }
    else
    {
        /*do nothing*/
    }

    if (u16Val > 0u)
    {
        sVn7x_u32ChnSts |= (uint32)1u << u8Chn;
    }
    else
    {
        sVn7x_u32ChnSts &= 0xFFFFFFFFul - ((uint32)1u << u8Chn);
    }
}
