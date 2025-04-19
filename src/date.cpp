#include <Rcpp.h>
#include <string>

// [[Rcpp::export]]
Rcpp::CharacterVector date_to_unix(Rcpp::CharacterVector x, std::string origin = "1900-01-01", bool datetime = false) {
  std::size_t n = x.size();
  Rcpp::CharacterVector result(n);

  double offset;
  if (origin == "1900-01-01") {
    offset = 25569.0;
  } else if (origin == "1904-01-01") {
    offset = 24107.0;
  } else if (origin == "1970-01-01") {
    offset = 0.0; // for hms
  } else {
    Rcpp::stop("Origin must be '1900-01-01' or '1904-01-01'");
  }

  for (std::size_t i = 0; i < n; ++i) {
    if (x[i] == NA_STRING) {
      result[i] = NA_STRING;
      continue;
    }

    std::string strVal = Rcpp::as<std::string>(x[i]);
    const char* str = strVal.c_str();
    char* endp = nullptr;
    double val = R_strtod(str, &endp);

    if (endp == str || !R_finite(val)) {
      result[i] = NA_STRING;
      continue;
    }

    if (!datetime && origin == "1900-01-01") {
      if (val < 60) {
        val += 1.0; // add one day
      }
    }

    double unixTime = (val - offset);
    if (datetime) {
      unixTime *= 86400.0; // for seconds
      if (origin != "1970-01-01") unixTime -= 7200.0;  // move two hours for UTC (2*60*60)
      if (origin == "1970-01-01") unixTime -= 3600.0;  // move one hour for hms
    }

    // Rf_PrintValue(Rcpp::wrap(unixTime));
    std::int64_t timestamp = static_cast<std::int64_t>(std::round(unixTime));

    result[i] = std::to_string(timestamp);
  }

  return result;
}
