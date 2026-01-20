#if !defined(PLATFORM_TYPES_H)
#define PLATFORM_TYPES_H
/**
 * \file
 *
 * \brief AUTOSAR Base
 *
 * This file contains the implementation of the AUTOSAR
 * module Base.
 *
 * \author Elektrobit Automotive GmbH, 91058 Erlangen, Germany
 *
 * Copyright 2005 - 2020 Elektrobit Automotive GmbH
 * All rights exclusively reserved for Elektrobit Automotive GmbH,
 * unless expressly agreed to otherwise.
 */

/*
 * MISRA-C:2012 Deviation List
 *
 * MISRAC2012-1) Deviated Rule: 20.10 (advisory)
 *   The # and ## preprocessor operators should not be used.
 *
 *   Reason:
 *   The operators are used to add the 'U' or 'UL' suffix to constants
 *   via macros which can only be performed using the '##' operator.
 */
/* !LINKSTO Base.ASR403.PLATFORM001,1, Base.ASR403.PLATFORM002,1, Base.ASR403.PLATFORM003,1 */

/*==================[macros]=================================================*/

/*------------------[AUTOSAR vendor identification]-------------------------*/

#if (defined PLATFORM_VENDOR_ID)
#error PLATFORM_VENDOR_ID is already defined
#endif
/** \brief AUTOSAR vendor identification: Elektrobit Automotive GmbH */
#define PLATFORM_VENDOR_ID         1U

/*------------------[AUTOSAR module identification]--------------------------*/

#if (defined PLATFORM_MODULE_ID)
#error PLATFORM_MODULE_ID already defined
#endif
/** \brief AUTOSAR module identification */
#define PLATFORM_MODULE_ID     199U

/*------------------[AUTOSAR release version identification]-----------------*/

#if (defined PLATFORM_AR_RELEASE_MAJOR_VERSION)
#error PLATFORM_AR_RELEASE_MAJOR_VERSION already defined
#endif
/** \brief AUTOSAR release major version */
#define PLATFORM_AR_RELEASE_MAJOR_VERSION     4U

#if (defined PLATFORM_AR_RELEASE_MINOR_VERSION)
#error PLATFORM_AR_RELEASE_MINOR_VERSION already defined
#endif
/** \brief AUTOSAR release minor version */
#define PLATFORM_AR_RELEASE_MINOR_VERSION     0U

#if (defined PLATFORM_AR_RELEASE_REVISION_VERSION)
#error PLATFORM_AR_RELEASE_REVISION_VERSION already defined
#endif
/** \brief AUTOSAR release revision version */
#define PLATFORM_AR_RELEASE_REVISION_VERSION  3U

/*------------------[AUTOSAR module version identification]-------------------*/

#if (defined PLATFORM_SW_MAJOR_VERSION)
#error PLATFORM_SW_MAJOR_VERSION already defined
#endif
/** \brief AUTOSAR module major version */
#define PLATFORM_SW_MAJOR_VERSION  2U

#if (defined PLATFORM_SW_MINOR_VERSION)
#error PLATFORM_SW_MINOR_VERSION already defined
#endif
/** \brief AUTOSAR module minor version */
#define PLATFORM_SW_MINOR_VERSION  5U

#if (defined PLATFORM_SW_PATCH_VERSION)
#error PLATFORM_SW_PATCH_VERSION already defined
#endif
/** \brief AUTOSAR module patch version */
#define PLATFORM_SW_PATCH_VERSION  0U

/*------------------[AUTOSAR module vendor API infix]------------------------*/

#if (defined PLATFORM_VENDOR_API_INFIX)
#error PLATFORM_VENDOR_API_INFIX already defined
#endif
/** \brief Vendor API infix
 *
 * Left empty as this header does not belong to any module */
#define PLATFORM_VENDOR_API_INFIX

/*------------------[global identifiers]-------------------------------------*/

/* !LINKSTO Base.ASR403.PLATFORM054,1 */
#if (!defined TRUE)
/** \brief true value for boolean type */
/* !LINKSTO Base.ASR403.PLATFORM056,1 */
#define TRUE 1U
#endif

/* !LINKSTO Base.ASR403.PLATFORM054,1 */
#if (!defined FALSE)
/** \brief false value for boolean type */
/* !LINKSTO Base.ASR403.PLATFORM056,1 */
#define FALSE 0U
#endif

/* redefinition check may lead to an incorrect condition that
 * TRUE and FALSE are equal */
/* !LINKSTO Base.ASR403.PLATFORM054,1 */
#if (TRUE == FALSE)
#error TRUE == FALSE is incorrect
#endif

/*------------------[CPU properties]-----------------------------------------*/

#if (defined CPU_TYPE_8)
#error CPU_TYPE_8 already defined
#endif
/** \brief cpu identifier for 8-bit CPUs */
/* !LINKSTO Base.ASR403.PLATFORM057,1, Base.ASR413.SWS_Platform_00064,1 */
#define CPU_TYPE_8 8U

#if (defined CPU_TYPE_16)
#error CPU_TYPE_16 already defined
#endif /* if (defined CPU_TYPE_16) */
/** \brief cpu identifier for 8-bit CPUs */
/* !LINKSTO Base.ASR403.PLATFORM057,1, Base.ASR413.SWS_Platform_00064,1 */
#define CPU_TYPE_16 16U

#if (defined CPU_TYPE_32)
#error CPU_TYPE_32 already defined
#endif
/** \brief cpu identifier for 32-bit CPUs */
/* !LINKSTO Base.ASR403.PLATFORM057,1, Base.ASR413.SWS_Platform_00064,1 */
#define CPU_TYPE_32 32U

/* !LINKSTO Base.ASR413.SWS_Platform_00064,1 */
#if (defined CPU_TYPE_64)
#error CPU_TYPE_64 already defined
#endif
/** \brief cpu identifier for 64-bit CPUs */
#define CPU_TYPE_64 64U

#if (defined MSB_FIRST)
#error MSB_FIRST already defined
#endif
/** \brief identifier for 'high byte first' */
#define MSB_FIRST 0U

#if (defined LSB_FIRST)
#error LSB_FIRST already defined
#endif /* if (defined LSB_FIRST) */
/** \brief identifier for 'little endian' */
#define LSB_FIRST 1U

#if (defined HIGH_BYTE_FIRST)
#error HIGH_BYTE_FIRST already defined
#endif
/** \brief identifier for 'high byte first' */
#define HIGH_BYTE_FIRST 0U

#if (defined LOW_BYTE_FIRST)
#error LOW_BYTE_FIRST already defined
#endif
/** \brief cpu identifier for  8-bit CPUs */
#define LOW_BYTE_FIRST 1U

#if (defined CPU_TYPE)
#error CPU_TYPE already defined
#endif

/** \brief Word size of this CPU (generated)
 *
 * Platforms Setting: 
 * The word size specifies the natural data unit size of a CPU. 
 * It is set in the Base configuration by
 * setting the parameter 'BaseCpuConfig_type'.
 *
 * Possible values:
 *  - CPU_TYPE_64 (if parameter 'BaseCpuConfig_type:64')
 *  - CPU_TYPE_32 (if parameter 'BaseCpuConfig_type:32')
 *  - CPU_TYPE_16 (if parameter 'BaseCpuConfig_type:16')
 *  - CPU_TYPE_8  (if parameter 'BaseCpuConfig_type:8')
 */
/* !LINKSTO Base.ASR403.PLATFORM044,1, Base.ASR403.PLATFORM045,1, Base.ASR413.SWS_Platform_00064,1, Base.ASR430.SWS_Platform_00045,1 */
#define CPU_TYPE CPU_TYPE_32

/* !LINKSTO Base.ASR403.PLATFORM043,1, Base.ASR403.PLATFORM048,1, Base.ASR403.PLATFORM049,1, Base.ASR403.PLATFORM038,1 */
#if (defined CPU_BIT_ORDER)
#error CPU_BIT_ORDER already defined
#endif

/** \brief Bit order of this CPU (generated)
 *
 * Platforms Setting:
 * The bit order of the CPU is set in the Base configuration 
 * by setting the parameter 'BaseCpuConfig_bitorder'. 
 *
 * Range:
 *  - LE (LSB_FIRST if parameter 'BaseCpuConfig_bitorder:LE')
 *  - BE (MSB_FIRST if parameter 'BaseCpuConfig_bitorder:BE')
 */
#define CPU_BIT_ORDER LSB_FIRST

/* !LINKSTO Base.ASR403.PLATFORM039,1, Base.ASR403.PLATFORM046,1, Base.ASR403.PLATFORM050,1, Base.ASR403.PLATFORM051,1 */
#if (defined CPU_BYTE_ORDER)
#error CPU_BYTE_ORDER already defined
#endif

/** \brief Byte order of this CPU (generated)
 *
 * Platforms Setting:
 * The byte order of the CPU is set in the Base configuration 
 * by setting the parameter 'BaseCpuConfig_byteorder'. 
 *
 * Range:
 *  - BE (HIGH_BYTE_FIRST if parameter 'BaseCpuConfig_byteorder:BE')
 *  - LE (LOW_BYTE_FIRST  if parameter 'BaseCpuConfig_byteorder:LE')
 */
#define CPU_BYTE_ORDER LOW_BYTE_FIRST

/* !LINKSTO Base.Dbg.ConfigMacro,1 */
#if (defined BASE_DBG_ENABLE)
#error BASE_DBG_ENABLE already defined
#endif

/** \brief Macro for enabling debug header file (generated)
 *
 * Range:
 *  - STD_ON (if configuration parameter 'Base/BaseDbg/BaseDbgHeaderFile' enabled)
 *  - STD_OFF (if configuration parameter 'Base/BaseDbg/BaseDbgHeaderFile' disabled)
 */
#define BASE_DBG_ENABLE STD_OFF
/* !LINKSTO Base.Properties.OverrideMemCpy,1 */
#if (defined TS_MEMCPY_CUSTOM_OVERRIDE)
#error TS_MEMCPY_CUSTOM_OVERRIDE already defined
#endif

/** \brief Macro for enabling overriding of the memory copy function by a custom implementation
 *
 * Range:
 *  - STD_ON (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemCpy' enabled)
 *  - STD_OFF (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemCpy' disabled)
 */
#define TS_MEMCPY_CUSTOM_OVERRIDE STD_OFF
/* !LINKSTO Base.Properties.OverrideMemSet,1 */
#if (defined TS_MEMSET_CUSTOM_OVERRIDE)
#error TS_MEMSET_CUSTOM_OVERRIDE already defined
#endif

/** \brief Macro for enabling overriding of the memory set function by a custom implementation
 *
 * Range:
 *  - STD_ON (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemSet' enabled)
 *  - STD_OFF (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemSet' disabled)
 */
#define TS_MEMSET_CUSTOM_OVERRIDE STD_OFF
/* !LINKSTO Base.Properties.OverrideMemCmp,1 */
#if (defined TS_MEMCMP_CUSTOM_OVERRIDE)
#error TS_MEMCMP_CUSTOM_OVERRIDE already defined
#endif

/** \brief Macro for enabling overriding of the memory compare function by a custom implementation
 *
 * Range:
 *  - STD_ON (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemCmp' enabled)
 *  - STD_OFF (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemCmp' disabled)
 */
#define TS_MEMCMP_CUSTOM_OVERRIDE STD_OFF
/* !LINKSTO Base.Properties.OverrideMemBZero,1 */
#if (defined TS_MEMBZERO_CUSTOM_OVERRIDE)
#error TS_MEMBZERO_CUSTOM_OVERRIDE already defined
#endif

/** \brief Macro for enabling overriding of the memory zeroing function by a custom implementation
 *
 * Range:
 *  - STD_ON (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemBZero' enabled)
 *  - STD_OFF (if configuration parameter 'Base/CustomOverrides/CustomOverride_MemBZero' disabled)
 */
#define TS_MEMBZERO_CUSTOM_OVERRIDE STD_OFF
/*==================[type definitions]=======================================*/

#if (!defined HAS_64BIT_TYPES)
#define HAS_64BIT_TYPES
#endif

/* !LINKSTO Base.ASR403.PLATFORM005,1, Base.ASR403.PLATFORM013,1, Base.ASR403.PLATFORM014,1, Base.ASR403.PLATFORM015,1, Base.ASR403.PLATFORM016,1, Base.ASR403.PLATFORM017,1, Base.ASR403.PLATFORM018,1, Base.ASR403.PLATFORM020,1, Base.ASR403.PLATFORM021,1, Base.ASR403.PLATFORM022,1, Base.ASR403.PLATFORM023,1, Base.ASR403.PLATFORM024,1, Base.ASR403.PLATFORM025,1, Base.ASR403.PLATFORM027,1, Base.ASR403.PLATFORM041,1, Base.ASR403.PLATFORM042,1, Base.ASR403.PLATFORM060,1, Base.ASR413.SWS_Platform_00066,1, Base.ASR413.SWS_Platform_00067,1 */
/** \brief Type definition of AUTOSAR standard type boolean
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.boolean.Mapping' 
 */
typedef unsigned char boolean;

/** \brief Type definition of AUTOSAR standard type sint8
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.sint8.Mapping' 
 */
typedef signed char sint8;

/** \brief Type definition of AUTOSAR standard type sint16
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.sint16.Mapping' 
 */
typedef signed short sint16;

/** \brief Type definition of AUTOSAR standard type sint32
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.sint32.Mapping' 
 */
typedef signed long sint32;

/** \brief Type definition of AUTOSAR standard type uint8
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.uint8.Mapping' 
 */
typedef unsigned char uint8;

/** \brief Type definition of AUTOSAR standard type uint16
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.uint16.Mapping' 
 */
typedef unsigned short uint16;

/** \brief Type definition of AUTOSAR standard type uint32
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.uint32.Mapping' 
 */
typedef unsigned long uint32;

/** \brief Type definition of AUTOSAR standard type sint8_least
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.sint8_least.Mapping' 
 */
typedef signed long sint8_least;

/** \brief Type definition of AUTOSAR standard type sint16_least
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.sint16_least.Mapping' 
 */
typedef signed long sint16_least;

/** \brief Type definition of AUTOSAR standard type sint32_least
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.sint32_least.Mapping' 
 */
typedef signed long sint32_least;

/** \brief Type definition of AUTOSAR standard type uint8_least
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.uint8_least.Mapping' 
 */
typedef unsigned long uint8_least;

/** \brief Type definition of AUTOSAR standard type uint16_least
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.uint16_least.Mapping' 
 */
typedef unsigned long uint16_least;

/** \brief Type definition of AUTOSAR standard type uint32_least
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.uint32_least.Mapping' 
 */
typedef unsigned long uint32_least;

/** \brief Type definition of AUTOSAR standard type float32
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.float32.Mapping' 
 */
typedef float float32;

/** \brief Type definition of AUTOSAR standard type float64
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.float64.Mapping' 
 */
typedef double float64;

/** \brief Type definition of AUTOSAR standard type sint64
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.sint64.Mapping' 
 */
typedef signed long long sint64;

/** \brief Type definition of AUTOSAR standard type uint64
 *
 * Platforms Setting: 
 * Definition of platform specific AUTOSAR standard types
 * that are set in the Platforms plugin by parameter 'Basetypes.uint64.Mapping' 
 */
typedef unsigned long long uint64;

/* !LINKSTO Base.Types.SizeType,1 */
/** \brief Type definition of platform specific size type (generated, depending on parameter 'Cpu.Type') */
typedef uint32 usize;

#if (!defined USIZE_C)
/** \brief Macro to define a constant of platform specific type usize (generated, depending on parameter 'Cpu.Type') */
/* Deviation MISRAC2012-1 <+2> */
#define USIZE_C(x) (x ## UL)
#endif

/* !LINKSTO Base.Types.VoidPtr,1 */
/** \brief type definition for pointer to void */
typedef void * VoidPtr;
/** \brief type definition for pointer to const void */
typedef const void * ConstVoidPtr;

/*==================[inclusions]=============================================*/

/* !LINKSTO Base.ConfigScheme.PlatformTypesGeneration,1 */
/** \brief Inclusions of custom standard files (generated)
 *
 * For each configuration parameter 'Base/CustomStdIncludeFiles/[file]',
 * an include statement is generated
 *
 * Platforms Setting: 
 * Custom standard file can be used for definition of
 * additional standard types that are customer specific.
 */

#endif /* if (!defined PLATFORM_TYPES_H) */
