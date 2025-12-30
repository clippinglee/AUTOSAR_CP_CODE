/*****************************************************************************************************************
******************************************************************************************************************
*  Copyright (C) .                                                                            
*  All rights reserved.                                                                                           
******************************************************************************************************************
*  FileName: BJT                                                                                             
*  Content:  BJT family drive
*  Category: 
******************************************************************************************************************
*  Revision Management                                                                                            
*  yyyy.mm.dd    name              version      description                                                       
*  ----------    --------          -------      -----------------------------------                               
*  2022.04.18    clipping            v0001        Frist edit                                                        
******************************************************************************************************************
******************************************************************************************************************/
/* Include Headerfiles  */
#ifndef _BJT_HWCFG_H_
#define _BJT_HWCFG_H_

#include "Bjt_Types.h"
#include "Pwm.h"
#include "Dio.h"


typedef enum
{
    BJT_ID_0,  /* OPH02  #0  */
    BJT_ID_1,  /* OPH01  #1  */

    BJT_ID_MAX
} Bjt_ChnIdType;

typedef struct 
{
    Bjt_ChnIdType eBjtChann;
    Bjt_InputChnType eBjt_Type;
    Pwm_ChannelType u8BjtPwmCntrl;
    Dio_ChannelType u8BjtDioCntrl;
    uint16 u16OLDiagAdcVal;
    uint16 u16ShortDiagAdcVal;
}Bjt_ChnCfgType;

#define BJT_ENABLE_PWM_TRIGGER_ADC
#define BJT_DISABLE_PEM_TRIGGER_ADC


extern const Bjt_ChnCfgType cBjt_atChannelInputCfg[BJT_ID_MAX];
extern const uint8 cBjt_au8AdcEidMap[BJT_ID_MAX];
#endif
