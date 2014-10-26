/**
*
*/

#pragma once

#include <chrono>
#include <boost/log/expressions.hpp>
#include <boost/log/sources/global_logger_storage.hpp>
#include <boost/log/support/date_time.hpp>
#include <boost/log/trivial.hpp>
#include <boost/log/utility/setup.hpp>

#define INFO_LOG  BOOST_LOG_SEV(my_logger::get(), boost::log::trivial::info)
#define WARN_LOG  BOOST_LOG_SEV(my_logger::get(), boost::log::trivial::warning)
#define ERROR_LOG BOOST_LOG_SEV(my_logger::get(), boost::log::trivial::error)

#define SYS_LOGFILE             "info.log"

//Narrow-char thread-safe logger.
typedef boost::log::sources::wseverity_logger_mt<boost::log::trivial::severity_level> logger_t;

//declares a global logger with a custom initialization
BOOST_LOG_GLOBAL_LOGGER(my_logger, logger_t)


inline std::chrono::steady_clock::time_point	TimerStart() {
	return std::chrono::steady_clock::now();
}

inline long long ProcessingTime(std::chrono::steady_clock::time_point start,
						std::chrono::steady_clock::time_point stop = std::chrono::steady_clock::now())
{
	return std::chrono::duration_cast<std::chrono::milliseconds>(stop - start).count();
}

inline std::wstring ProcessingTimeStr(std::chrono::steady_clock::time_point start,
	std::chrono::steady_clock::time_point stop = std::chrono::steady_clock::now())
{
	std::wstringstream ss;
	ss << ProcessingTime(start, stop) << L"ms";
	return ss.str();
}















