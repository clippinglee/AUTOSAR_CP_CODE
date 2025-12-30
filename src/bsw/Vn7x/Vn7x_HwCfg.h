

#ifndef _VN7X_HWCFG_H_
#define _VN7X_HWCFG_H_

#include "Vn7x_Types.h"
#include "Pwm.h"
#include "Dio.h"


typedef enum
{
    VN7X_ID_VN7E010AJ_0,  /* HSDC_OUT0 */
    VN7X_ID_VN7E010AJ_1,  /* HSDC_OUT1 */
    VN7X_ID_VN7E010AJ_2,  /* HSDC_OUT2 */
    VN7X_ID_VN7E010AJ_3,  /* HSDC_OUT3 */
    VN7X_ID_VN7E010AJ_4,  /* HSDC_OUT4 */
    VN7X_ID_VN7E010AJ_5,  /* HSDC_OUT5 */
    VN7X_ID_VN7E010AJ_6,  /* HSDC_OUT6 */
    VN7X_ID_VN7E010AJ_7,  /* HSDC_OUT7 */

    VN7X_ID_MAX
} Vn7x_ChnIdType;

typedef struct 
{
    Vn7x_ChnIdType eVn7xChann;
    Vn7x_InputChnType eVn7x_Type;
    Pwm_ChannelType u8Vn7xPwmCntrl;
    Dio_ChannelType u8Vn7xDioInput;
    Dio_ChannelType u8Vn7xDioSEn;
    Dio_ChannelType u8Vn7xDioFaultRST;
    uint16 u16OLDiagAdcVal;
    uint16 u16ShortDiagAdcVal;
}Vn7x_ChnCfgType;


#define VN7X_DAIG_SEL_CHN_MAX 0u
#define VN7X_DAIG_SEL_CHN_ZERO 0u

#define VN7X_ENABLE_PWM_TRIGGER_ADC
#define VN7X_DISABLE_PEM_TRIGGER_ADC


extern const Vn7x_ChnCfgType cVn7x_atChannelInputCfg[VN7X_ID_MAX];
extern const uint8 cVn7x_au8AdcEidMap[VN7X_ID_MAX];
#endif
