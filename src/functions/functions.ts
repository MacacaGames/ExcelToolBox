/* global clearInterval, console, CustomFunctions, setInterval */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

// /**
//  * convert data to Json
//  * @customfunction
//  * @param headers Json element name
//  * @param datas Json element value
//  * @returns {string} The Json of the data.
//  */
// export function tojson(headers: string, datas: string): string {
//   const objects = getJsonArrayFromData(headers, datas);
//   const json = JSON.stringify(objects);
//   if (json.length < 50000) {
//     return json;
//   }
//   return json;
// }

/**
 * convert data to Json
 * @customfunction
 * @param datas element value
 * @returns {string} The Json of the data.
 */
export function tojson(datas: string): string {
  const rows = datas.trim().split('\n');
  const array = rows.map(row => row.split(','));
  return array.toString();
}

function dataArrayToJson(data: any[][]): string {
  const headers = data[0];
  const rows = data.slice(1);

  const json = rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      let value = row[index];
      if(value === "NoData_Int" || value === "NoData_Enum") {
        value = 0;
      } else if (value === "NoData_Bool") {
        value = false;
      } else if (value === "NoData_Json") {
        value = {};
      } else if (value === "NoData_Array") {
        value = [];
      } else if (value === "NoData_String") {
        value = "";
      }
      obj[header] = value;
    });
    return obj;
  });

  return JSON.stringify(json, null, 2);
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
