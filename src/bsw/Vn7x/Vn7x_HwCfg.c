

#include "Vn7x_HwCfg.h"
#include "Dio.h"



const Vn7x_ChnCfgType cVn7x_atChannelInputCfg[VN7X_ID_MAX] = 
{
    {
        VN7X_ID_VN7E010AJ_0,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P01_02,
        DioConf_DioChannel_DioChannel_P01_01,
        DioConf_DioChannel_DioChannel_P01_00,
        0xFFF,
        0xFFF
    },
    {
        VN7X_ID_VN7E010AJ_1,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P01_09,
        DioConf_DioChannel_DioChannel_P01_10,
        DioConf_DioChannel_DioChannel_P01_08,
        0xFFF,
        0xFFF
    },
    {
        VN7X_ID_VN7E010AJ_2,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P01_12,
        DioConf_DioChannel_DioChannel_P01_13,
        DioConf_DioChannel_DioChannel_P01_11,
        0xFFF,
        0xFFF
    },
    {
        VN7X_ID_VN7E010AJ_3,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P00_14,
        DioConf_DioChannel_DioChannel_P00_15,
        DioConf_DioChannel_DioChannel_P00_13,
        0xFFF,
        0xFFF
    },
    {
        VN7X_ID_VN7E010AJ_4,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P00_01,
        DioConf_DioChannel_DioChannel_P02_08,
        DioConf_DioChannel_DioChannel_P00_00,
        0xFFF,
        0xFFF
    },
    {
        VN7X_ID_VN7E010AJ_5,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P02_11,
        DioConf_DioChannel_DioChannel_P02_10,
        DioConf_DioChannel_DioChannel_P02_09,
        0xFFF,
        0xFFF
    },
    {
        VN7X_ID_VN7E010AJ_6,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P13_03,
        DioConf_DioChannel_DioChannel_P13_04,
        DioConf_DioChannel_DioChannel_P02_12,
        0xFFF,
        0xFFF
    },
    {
        VN7X_ID_VN7E010AJ_7,
        VN7X_DIO,
        VN7X_DIO_PWM_INVALIDVAL,
        DioConf_DioChannel_DioChannel_P14_12,
        DioConf_DioChannel_DioChannel_P14_13,
        DioConf_DioChannel_DioChannel_P02_14,
        0xFFF,
        0xFFF
    },
};

const uint8 cVn7x_au8AdcEidMap[VN7X_ID_MAX] =
{
    1,2,3,4,5,6
};

