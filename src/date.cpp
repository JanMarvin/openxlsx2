#include <string>
#include "openxlsx2.h"

// [[Rcpp::export]]
Rcpp::CharacterVector date_to_unix(Rcpp::CharacterVector x, std::string origin = "1900-01-01", bool datetime = true) {
  R_xlen_t n = x.size();
  Rcpp::CharacterVector result(n);

  double epoch_offset_days; // Days from the specified origin epoch to 1970-01-01

  if (origin == "1900-01-01") {
    // Excel 1900 system: epoch is effectively 1899-12-30, Day 0
    // 1900-01-01 is Day 2.
    // Unix epoch 1970-01-01 is 25569 days after 1899-12-30.
    epoch_offset_days = 25569.0;
  } else if (origin == "1904-01-01") {
    // Excel 1904 system: epoch is 1904-01-01, Day 0
    // Unix epoch 1970-01-01 is 24107 days after 1904-01-01.
    epoch_offset_days = 24107.0;
  } else if (origin == "1970-01-01") {
    // Assuming 1970 origin means days since 1970-01-01 (like Unix epoch in days)
    epoch_offset_days = 0.0;
  }
  else {
    Rcpp::stop("origin_system must be '1900', '1904', or '1970'.");
  }


  for (R_xlen_t i = 0; i < n; ++i) {

    char* endp = nullptr;
    double val = R_strtod(x[i], &endp);

    if (endp == x[i] || !R_finite(val)) {
      // result[i] = NA_STRING;
      continue;
    }

    // Handle the Excel 1900 leap year bug for dates before March 1, 1900
    // This is applied only if the origin system is 1900
    if (origin == "1900-01-01" && val < 60) {
      // If val is 60 (1900-02-29 in Excel), it's an invalid date.
      // For val < 60, Excel's date is off by 1 day. Add 1 to correct.
      // Note: Handling of the invalid 60 itself might need specific logic
      // depending on desired behavior (e.g., return NA).
      if (val > 0 && val < 60) { // Apply correction for dates between 1900-01-01 (serial 1) and 1900-02-28 (serial 59)
        val += 1.0;
      } else if (val == 60) {
        // Serial 60 corresponds to 1900-02-29, which is invalid.
        // Returning NA or a specific value might be appropriate.
        // For now, let's proceed with the value, which will map to 1900-02-29
        // in the standard conversion, but be aware this is an Excel quirk.
      }
    }


    double total_days_since_excel_epoch = val;
    double days_since_unix_epoch = total_days_since_excel_epoch - epoch_offset_days;

    double unix_timestamp; // in seconds since 1970-01-01 UTC

    if (datetime) {
      unix_timestamp = days_since_unix_epoch * 86400.0; // Convert days to seconds
    } else {
      // If include_time is false, we only care about the date part at midnight UTC
      double date_part_val = std::floor(val);
      double days_since_unix_epoch_date_part = date_part_val - epoch_offset_days;
      unix_timestamp = days_since_unix_epoch_date_part;
    }

    // Round to the nearest second for the timestamp
    result[i] = std::to_string(std::round(unix_timestamp));
  }

  return result;
}
