/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import { db, handleFirestoreError, OperationType, testConnection } from './firebase';
import {
  collection,
  doc,
  setDoc,
  updateDoc,
  deleteDoc,
  onSnapshot,
  query,
  orderBy,
  limit,
  getDocs
} from 'firebase/firestore';

// Initial default menu
const DEFAULT_MENU: { c: string; i: [string, number, string, boolean, string][] }[] = [
  {
    c: "Mains",
    i: [
      ["Cheeseburger", 12.5, "🍔", false, ""],
      ["Chicken sandwich", 11, "🥪", false, ""],
      ["Margherita pizza", 13, "🍕", true, "Fresh mozzarella and basil on a hand-stretched crust."],
      ["Pad thai", 14, "🍜", false, ""],
      ["Caesar salad", 10.5, "🥗", false, ""],
      ["Fish and chips", 15, "🐟", false, ""]
    ]
  },
  {
    c: "Sides",
    i: [
      ["Fries", 4.5, "🍟", false, ""],
      ["Onion rings", 5.5, "🧅", false, ""],
      ["Side salad", 4, "🥬", false, ""],
      ["Soup of the day", 6, "🍲", false, ""],
      ["Iced Latte", 3.5, "☕", false, ""]
    ]
  },
  {
    c: "Drinks",
    i: [
      ["Soda", 3, "🥤", false, ""],
      ["Iced tea", 3.5, "🧊", false, ""],
      ["Coffee", 3, "☕", false, ""],
      ["Lemonade", 3.5, "🍋", true, "Freshly squeezed and served over ice."],
      ["Beer", 6.5, "🍺", false, ""],
      ["House wine", 8, "🍷", false, ""]
    ]
  },
  {
    c: "Desserts",
    i: [
      ["Cheesecake", 7, "🍰", false, ""],
      ["Brownie", 6, "🍫", false, ""],
      ["Ice cream", 5, "🍨", false, ""],
    ]
  }
];

const TAX = 0.08;
const TABLES = ["1", "2", "3", "4", "5", "6", "7", "8", "Takeout"];
const CODE_CH = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";

const KH_DICT: Record<string, string> = {
  "Add item": "បន្ថែមមុខម្ហូប",
  "Advertisement picture for the pop-up (one picture)": "រូបភាពផ្សព្វផ្សាយសម្រាប់ផ្ទាំងលេចឡើង (មួយសន្លឹក)",
  "All saved orders, newest guest first. Tap a number to expand.": "ការកម្ម៉ង់ទាំងអស់ដែលបានរក្សាទុក ភ្ញៀវថ្មីបំផុតមុន។ ចុចលើលេខដើម្បីពង្រីក។",
  "Back": "ត្រឡប់ក្រោយ",
  "Back to POS": "ត្រឡប់ទៅប្រព័ន្ធលក់",
  "Bot token (from @BotFather)": "Bot token (ពី @BotFather)",
  "Card": "កាត",
  "Cash": "សាច់ប្រាក់",
  "Cash received": "សាច់ប្រាក់ទទួលបាន",
  "Cash received - dollars ($)": "សាច់ប្រាក់ទទួលបាន - ដុល្លារ ($)",
  "Cash received - riel (៛)": "សាច់ប្រាក់ទទួលបាន - រៀល (៛)",
  "Chat ID (a number, or @channelname)": "Chat ID (លេខ ឬ @channelname)",
  "Clear all sales data": "លុបទិន្នន័យលក់ទាំងអស់",
  "Clear order": "លុបការកម្ម៉ង់",
  "Close": "បិទ",
  "Complete sale": "បញ្ចប់ការលក់",
  "Dashboard": "របាយការណ៍លក់",
  "Database": "មូលដ្ឋានទិន្នន័យ",
  "Done": "រួចរាល់",
  "Download backup (JSON)": "ទាញយកការបម្រុងទុក (JSON)",
  "Due": "ត្រូវបង់",
  "Due in riel": "ត្រូវបង់ជារៀល",
  "Edit menu": "កែម៉ឺនុយ",
  "Exchange rate (riel per 1 dollar)": "អត្រាប្តូរប្រាក់ (រៀលក្នុង ១ ដុល្លារ)",
  "Export Excel": "នាំចេញ Excel",
  "Export PDF": "នាំចេញ PDF",
  "Guest order history by telephone": "ប្រវត្តិកម្ម៉ង់របស់ភ្ញៀវតាមលេខទូរស័ព្ទ",
  "Guest telephone (9 or 10 digits)": "លេខទូរស័ព្ទភ្ញៀវ (៩ ឬ ១០ ខ្ទង់)",
  "Payment methods": "វិធីទូទាត់",
  "Photo (optional)": "រូបថត (ស្រេចចិត្ត)",
  "Promote this item in the pop-up": "ផ្សព្វផ្សាយមុខម្ហូបនេះក្នុងផ្ទាំងលេចឡើង",
  "Promotion categories. Add any item to the cart for the open table.": "ប្រភេទប្រូម៉ូសិន។ ដាក់មុខម្ហូបណាមួយក្នុងកន្ត្រកសម្រាប់តុដែលកំពុងបើក។",
  "Promotion headline": "ចំណងជើងប្រូម៉ូសិន",
  "Promotion page link (optional, starts with https://)": "តំណទំព័រប្រូម៉ូសិន (ស្រេចចិត្ត ចាប់ផ្តើមដោយ https://)",
  "Promotions": "ប្រូម៉ូសិន",
  "Receipt": "បង្កាន់ដៃ",
  "Recent orders": "ការកម្ម៉ង់ថ្មីៗ",
  "Remove ad picture": "លុបរូបភាពផ្សព្វផ្សាយ",
  "Revenue": "ចំណូល",
  "Revenue (KHR)": "ចំណូល (រៀល)",
  "Sales dashboard": "ផ្ទាំងរបាយការណ៍លក់",
  "Save": "រក្សាទុក",
  "Send test message": "ផ្ញើសារសាកល្បង",
  "Send test order alert": "ផ្ញើដំណឹងសាកល្បងនៃការកម្ម៉ង់",
  "Alert to Telegram admin group when guest orders": "ផ្ញើដំណឹងទៅក្រុម Telegram នៅពេលមានការកម្ម៉ង់ពីភ្ញៀវ",
  "Send the receipt automatically after each payment": "ផ្ញើបង្កាន់ដៃដោយស្វ័យប្រវត្តិបន្ទាប់ពីការទូទាត់នីមួយៗ",
  "Send to Telegram": "ផ្ញើទៅ Telegram",
  "Subtotal": "សរុបរង",
  "Take payment": "ទូទាត់ប្រាក់",
  "Tax (8%)": "ពន្ធ (8%)",
  "Theme": "ពណ៌ផ្ទៃ",
  "Tip": "ប្រាក់ទិប",
  "Top items": "មុខម្ហូបលក់ដាច់",
  "Total": "សរុប",
  "Total in riel": "សរុបជារៀល",
  "Grand total": "សរុបរួម",
  "Grand total in riel": "សរុបរួមជារៀល",
  "Upload this device's saved data": "បញ្ចូលទិន្នន័យដែលបានរក្សាទុកក្នុងឧបករណ៍នេះ",
  "Category": "ប្រភេទ",
  "Item name": "ឈ្មោះមុខម្ហូប",
  "Price": "តម្លៃ",
  "Price, e.g. 9.50": "តម្លៃ ឧ. 9.50",
  "Promo description": "ការពិពណ៌នាប្រូម៉ូសិន",
  "Promo description (optional)": "ការពិពណ៌នាប្រូម៉ូសិន (ស្រេចចិត្ត)",
  "Restaurant name": "ឈ្មោះភោជនីយដ្ឋាន",
  "Search telephone": "ស្វែងរកលេខទូរស័ព្ទ",
  "Tables": "តុ",
  "Food promotions": "ប្រូម៉ូសិនម្ហូប",
  "Leave empty to use the built-in promotions page": "ទុកទទេដើម្បីប្រើទំព័រប្រូម៉ូសិនដែលមានស្រាប់",
  "Takeout": "ខ្ចប់ទៅ",
  "No items yet. Tap Add to cart on a menu item.": "មិនទាន់មានមុខម្ហូបទេ។ ចុច «ដាក់ក្នុងកន្ត្រក» លើមុខម្ហូប។",
  "+ Add to cart": "+ ដាក់ក្នុងកន្ត្រក",
  "Added - tap to add another": "បានបន្ថែម - ចុចដើម្បីបន្ថែមទៀត",
  "★ Promotions": "★ ប្រូម៉ូសិន",
  "★ Promoted": "★ កំពុងផ្សព្វផ្សាយ",
  "Promote": "ផ្សព្វផ្សាយ",
  "Promote category": "ផ្សព្វផ្សាយប្រភេទ",
  "Unpromote category": "ឈប់ផ្សព្វផ្សាយប្រភេទ",
  "Photo": "រូបថត",
  "Remove": "លុប",
  "No tip": "គ្មានប្រាក់ទិប",
  "Today": "ថ្ងៃនេះ",
  "Last 7 days": "៧ ថ្ងៃចុងក្រោយ",
  "All time": "គ្រប់ពេល",
  "Orders": "ការកម្ម៉ង់",
  "Average order": "ការកម្ម៉ង់ជាមធ្យម",
  "Tips": "ប្រាក់ទិប",
  "Revenue by hour": "ចំណូលតាមម៉ោង",
  "Revenue by day": "ចំណូលតាមថ្ងៃ",
  "Date": "កាលបរិច្ឆេទ",
  "Time": "ម៉ោង",
  "Table": "តុ",
  "Pay": "ទូទាត់",
  "Phone": "ទូរស័ព្ទ",
  "Items": "មុខម្ហូប",
  "Item": "មុខម្ហូប",
  "Qty": "ចំនួន",
  "No sales in this range yet.": "មិនទាន់មានការលក់ក្នុងអំឡុងពេលនេះទេ។",
  "No data yet.": "មិនទាន់មានទិន្នន័យទេ។",
  "No orders in this range yet.": "មិនទាន់មានការកម្ម៉ង់ក្នុងអំឡុងពេលនេះទេ។",
  "No orders with a telephone number yet.": "មិនទាន់មានការកម្ម៉ង់ដែលមានលេខទូរស័ព្ទទេ។",
  "No guest matches that number.": "រកមិនឃើញភ្ញៀវដែលមានលេខនេះទេ។",
  "Enter the guest telephone: 9 or 10 digits, numbers only.": "សូមបញ្ចូលលេខទូរស័ព្ទភ្ញៀវ ៩ ឬ ១០ ខ្ទង់ (លេខតែប៉ុណ្ណោះ)។",
  "Card ": "កាត",
  "Mains": "មុខម្ហូបចម្បង",
  "Sides": "មុខម្ហូបបន្ថែម",
  "Drinks": "ភេសជ្ជៈ",
  "Desserts": "បង្អែម",
  "Cheeseburger": "ឆីសប៊ឺហ្គឺ",
  "Chicken sandwich": "សាំងវិចមាន់",
  "Margherita pizza": "ភីហ្សាម៉ាហ្គារីតា",
  "Pad thai": "ផាត់ថៃ",
  "Caesar salad": "សាឡាតសេសារ",
  "Fish and chips": "ត្រីបំពង និងដំឡូងបំពង",
  "Fries": "ដំឡូងបំពង",
  "Onion rings": "ខ្ទឹមបារាំងបំពង",
  "Side salad": "សាឡាតបន្ថែម",
  "Soup of the day": "ស៊ុបប្រចាំថ្ងៃ",
  "Iced Latte": "កាហ្វេឡាតេទឹកកក",
  "Soda": "ទឹកកាបូនិក",
  "Iced tea": "តែទឹកកក",
  "Coffee": "កាហ្វេ",
  "Lemonade": "ទឹកក្រូចឆ្មារ",
  "Beer": "ស្រាបៀរ",
  "House wine": "ស្រាទំពាំងបាយជូរ",
  "Cheesecake": "ឆីសខេក",
  "Brownie": "ប្រោនី",
  "Ice cream": "ការ៉េម",
  "Today's specials": "ពិសេសប្រចាំថ្ងៃ",
  "Choose a page": "ជ្រើសរើសទំព័រ",
  "Guest order": "ភ្ញៀវកម្ម៉ង់",
  "Order from the table": "កម្ម៉ង់ពីតុ",
  "Staff": "បុគ្គលិក",
  "Register and guest orders": "ការលក់ និងការកម្ម៉ង់ភ្ញៀវ",
  "Admin": "អ្នកគ្រប់គ្រង",
  "Menu, reports and settings": "ម៉ឺនុយ របាយការណ៍ និងការកំណត់",
  "Enter PIN": "បញ្ចូលលេខសម្ងាត់ PIN",
  "OK": "យល់ព្រម",
  "Cancel": "បោះបង់",
  "Incorrect PIN.": "PIN មិនត្រឹមត្រូវ។",
  "Switch page": "ប្តូរទំព័រ",
  "Guest orders": "ការកម្ម៉ង់ពីភ្ញៀវ",
  "Access": "ការចូលប្រើ",
  "Staff login": "ចូលជាបុគ្គលិក",
  "Your table": "តុរបស់អ្នក",
  "View promotion page": "មើលទំព័រប្រូម៉ូសិន",
  "Visit promotion website ↗": "ចូលមើលគេហទំព័រប្រូម៉ូសិន ↗",
  "View order": "មើលការកម្ម៉ង់",
  "Your order": "ការកម្ម៉ង់របស់អ្នក",
  "Place order": "បញ្ជូនការកម្ម៉ង់",
  "Telephone (optional)": "លេខទូរស័ព្ទ (ស្រេចចិត្ត)",
  "Telegram username (optional)": "ឈ្មោះគណនី Telegram (ស្រេចចិត្ត)",
  "Telegram username, e.g. @username (optional)": "ឈ្មោះគណនី Telegram ឧ. @username (ស្រេចចិត្ត)",
  "Note for the kitchen (optional)": "កំណត់ចំណាំសម្រាប់ផ្ទះបាយ (ស្រេចចិត្ត)",
  "Thank you! Your order was sent to the staff.": "សូមអរគុណ! ការកម្ម៉ង់របស់អ្នកត្រូវបានផ្ញើទៅបុគ្គលិក។",
  "New order": "កម្ម៉ង់ថ្មី",
  "Choose your table first.": "សូមជ្រើសរើសតុរបស់អ្នកជាមុនសិន។",
  "No new guest orders.": "គ្មានការកម្ម៉ង់ថ្មីពីភ្ញៀវទេ។",
  "Dismiss": "មិនទទួល",
  "Add to table cart": "ដាក់ក្នុងកន្ត្រកតុ",
  "Admin PIN": "PIN អ្នកគ្រប់គ្រង",
  "Staff PIN": "PIN បុគ្គលិក",
  "Remove PIN": "លុប PIN",
  "No items yet.": "មិនទាន់មានមុខម្ហូបទេ។",
  "Order": "ការកម្ម៉ង់",
  "Payment": "ការទូទាត់",
  "Cash on Delivery": "បង់ប្រាក់ពេលទទួលទំនិញ",
  "Upload the payment screenshot (required)": "ផ្ទុកឡើងរូបថតអេក្រង់នៃការទូទាត់ (ចាំបាច់)",
  "Screenshot ready.": "រូបថតអេក្រង់រួចរាល់។",
  "Please upload your payment screenshot.": "សូមផ្ទុកឡើងរូបថតអេក្រង់នៃការទូទាត់របស់អ្នក។",
  "KHQR is not available yet. Please choose Cash on Delivery.": "KHQR មិនទាន់អាចប្រើបានទេ។ សូមជ្រើសរើសបង់ប្រាក់ពេលទទួលទំនិញ។",
  "Payment matches": "ការទូទាត់ត្រូវគ្នា",
  "Not matching": "មិនត្រូវគ្នា",
  "Payment matched": "ការទូទាត់ត្រូវគ្នា",
  "Waiting for your check": "រង់ចាំការពិនិត្យរបស់អ្នក",
  "Check the KHQR payment first.": "សូមពិនិត្យការទូទាត់ KHQR ជាមុនសិន។",
  "Write a reason first.": "សូមសរសេរមូលហេតុជាមុនសិន។",
  "KHQR payment matched. Thank you!": "ការទូទាត់ KHQR ត្រឹមត្រូវ។ សូមអរគុណ!",
  "KHQR payment: staff are checking your screenshot.": "ការទូទាត់ KHQR៖ បុគ្គលិកកំពុងពិនិត្យរូបថតអេក្រង់របស់អ្នក។",
  "Pay with cash when your order is delivered.": "បង់ប្រាក់សាច់ប្រាក់នៅពេលទទួលការកម្ម៉ង់។",
  "Upload new screenshot": "ផ្ទុកឡើងរូបថតអេក្រង់ថ្មី",
  "Pay by Cash on Delivery": "បង់ប្រាក់ពេលទទួលទំនិញ",
  "Send": "ផ្ញើ",
  "KHQR payment code image (for guests to scan)": "រូបភាពកូដទូទាត់ KHQR (សម្រាប់ភ្ញៀវស្កេន)",
  "KHQR account name (shown to guests)": "ឈ្មោះគណនី KHQR (បង្ហាញដល់ភ្ញៀវ)",
  "Remove KHQR image": "លុបរូបភាព KHQR",
  "Marked as not matching": "បានកំណត់ថាមិនត្រូវគ្នា",
  "Payment marked as matching": "បានកំណត់ថាការទូទាត់ត្រូវគ្នា",
  "Your order code": "លេខកូដកម្ម៉ង់របស់អ្នក",
  "Assign to table": "កំណត់ទៅតុ",
  "Your receipt": "បង្កាន់ដៃរបស់អ្នក",
  "View receipt": "មើលបង្កាន់ដៃ",
  "Send by SMS": "ផ្ញើតាម SMS",
  "Share or copy": "ចែករំលែក ឬចម្លង",
  "Confirm": "បញ្ជាក់",
  "Confirmed": "បានបញ្ជាក់",
  "Decline": "បដិសេធ",
  "Declined": "បានបដិសេធ",
  "Recent decisions": "ការសម្រេចថ្មីៗ",
  "Note / reason (required to decline)": "កំណត់ចំណាំ / មូលហេតុ (ត្រូវការនៅពេលបដិសេធ)",
  "Write a reason to decline.": "សូមសរសេរមូលហេតុនៃការបដិសេធ។"
};

const KH_RX: [RegExp, (m: RegExpMatchArray) => string][] = [
  [/^KHQR payment is not matching\.(?: Reason: (.+))?$/, m => "ការទូទាត់ KHQR មិនត្រូវគ្នាទេ។" + (m[1] ? " មូលហេតុ: " + m[1] : "")],
  [/^Not matching(?:: (.+))?$/, m => "មិនត្រូវគ្នា" + (m[1] ? ": " + m[1] : "")],
  [/^Scan with your banking app and pay exactly (.+), then upload the screenshot\.$/, m => "ស្កេនជាមួយកម្មវិធីធនាគារ ហើយបង់ឱ្យបានត្រឹមត្រូវ " + m[1] + " បន្ទាប់មកផ្ទុកឡើងរូបថតអេក្រង់។"],
  [/^Order (#\w+) confirmed\.(?: (?:Table )?(\S+)\.)?(?: Note: (.+))?$/, m => "ការកម្ម៉ង់ " + m[1] + " ត្រូវបានបញ្ជាក់។" + (m[2] ? (/^\d+$/.test(m[2]) ? " តុ " + m[2] + "។" : " ខ្ចប់ទៅ។") : "") + (m[3] ? " កំណត់ចំណាំ: " + m[3] : "")],
  [/^Your order has been confirmed by the staff\.(?: (?:Table )?(\S+)\.)?(?: Note: (.+))?$/, m => "បុគ្គលិកបានបញ្ជាក់ការកម្ម៉ង់របស់អ្នក។" + (m[1] ? (/^\d+$/.test(m[1]) ? " តុ " + m[1] + "។" : " ខ្ចប់ទៅ។") : "") + (m[2] ? " កំណត់ចំណាំ: " + m[2] : "")],
  [/^Order (#\w+)$/, m => "ការកម្ម៉ង់ " + m[1]],
  [/^Order (#\w+) confirmed$/, m => "ការកម្ម៉ង់ " + m[1] + " ត្រូវបានបញ្ជាក់"],
  [/^Order (#\w+) declined$/, m => "ការកម្ម៉ង់ " + m[1] + " ត្រូវបានបដិសេធ"],
  [/^Your order has been confirmed by the staff\.(?: Note: (.+))?$/, m => "បុគ្គលិកបានបញ្ជាក់ការកម្ម៉ង់របស់អ្នក។" + (m[1] ? " កំណត់ចំណាំ: " + m[1] : "")],
  [/^Sorry, your order was declined\.(?: Reason: (.+))?$/, m => "សូមអភ័យទោស ការកម្ម៉ង់របស់អ្នកត្រូវបានបដិសេធ។" + (m[1] ? " មូលហេតុ: " + m[1] : "")],
  [/^Order (#\w+) confirmed\.(?: Note: (.+))?$/, m => "ការកម្ម៉ង់ " + m[1] + " ត្រូវបានបញ្ជាក់។" + (m[2] ? " កំណត់ចំណាំ: " + m[2] : "")],
  [/^Order (#\w+) was declined\.(?: Reason: (.+))?$/, m => "ការកម្ម៉ង់ " + m[1] + " ត្រូវបានបដិសេធ។" + (m[2] ? " មូលហេតុ: " + m[2] : "")],
  [/^Order (#\w+) sent\. Waiting for staff to confirm\.$/, m => "បានផ្ញើការកម្ម៉ង់ " + m[1] + "។ កំពុងរង់ចាំបុគ្គលិកបញ្ជាក់។"],
  [/^(\d+) items? - (.+)$/, m => `${m[1]} មុខ - ${m[2]}`],
  [/^Table (\d+)$/, m => `តុ ${m[1]}`],
  [/^Cart - Table (\d+)$/, m => `កន្ត្រក - តុ ${m[1]}`],
  [/^Cart - Takeout$/, () => "កន្ត្រក - ខ្ចប់ទៅ"],
  [/^(\d+) sold$/, m => `${m[1]} បានលក់`],
  [/^Change: (.+)$/, m => `ប្រាក់អាប់: ${m[1]}`],
  [/^Short by (.+)$/, m => `ខ្វះ ${m[1]}`],
  [/^Enter an amount of at least (.+)$/, m => `សូមបញ្ចូលចំនួនយ៉ាងតិច ${m[1]}`],
  [/^Rate: \$1 = ៛(.+)$/, m => `អត្រា: $1 = ៛${m[1]}`],
  [/^(\d+) orders? - (.+) - last visit (.+)$/, m => `${m[1]} ការកម្ម៉ង់ - ${m[2]} - ចុងក្រោយ ${m[3]}`],
  [/^(Table \d+|Takeout) - (.+) - paid by (Card|Cash|KHQR)(?: - Tel (.+))?$/, m => (/^Table/.test(m[1]) ? "តុ " + m[1].slice(6) : "ខ្ចប់ទៅ") + " - " + m[2] + " - បង់តាម" + (m[3] === "Card" ? "កាត" : m[3] === "KHQR" ? "KHQR" : "សាច់ប្រាក់") + (m[4] ? " - ទូរស័ព្ទ " + m[4] : "")]
];

interface SaleRecord {
  id?: string;
  table: string;
  method: string;
  tel: string;
  total: number;
  tip: number;
  sub: number;
  tax: number;
  rate: number;
  date: string;
  ts: number;
  items: [string, number, number][];
  time: string;
}

interface GuestOrder {
  id: string;
  code: string;
  table: string;
  pay: string;
  payStatus: string;
  payNote?: string;
  proof?: string;
  items: [string, number, number][];
  total: number;
  rate: number;
  tel?: string;
  tgUser?: string;
  note?: string;
  ts: number;
  date: string;
  time: string;
  status: 'new' | 'confirmed' | 'declined';
  reason?: string;
  decidedAt?: number;
  receipt?: any;
  discount?: number;
  discountPercent?: number;
  tgMsgId?: number;
}

// Convert data URL to Blob for Telegram sendPhoto
function dataURItoBlob(dataURI: string): Blob {
  const byteString = atob(dataURI.split(',')[1]);
  const mimeString = dataURI.split(',')[0].split(':')[1].split(';')[0];
  const ab = new ArrayBuffer(byteString.length);
  const ia = new Uint8Array(ab);
  for (let i = 0; i < byteString.length; i++) {
    ia[i] = byteString.charCodeAt(i);
  }
  return new Blob([ab], { type: mimeString });
}

// Format date to YYYY-MM-DD HH:mm:ss
function formatTgDate(date: Date = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  const Y = date.getFullYear();
  const M = pad(date.getMonth() + 1);
  const D = pad(date.getDate());
  const h = pad(date.getHours());
  const m = pad(date.getMinutes());
  const s = pad(date.getSeconds());
  return `${Y}-${M}-${D} ${h}:${m}:${s}`;
}

export default function App() {
  // Persistence & Config
  const [role, setRole] = useState<string>(() => sessionStorage.getItem('pos-role') || '');
  const [rname, setRname] = useState<string>(() => localStorage.getItem('pos-name') || 'Corner Kitchen');
  const [lang, setLang] = useState<string>(() => localStorage.getItem('pos-lang') === 'kh' ? 'kh' : 'en');
  const [rate, setRate] = useState<number>(() => {
    const r = Number(localStorage.getItem('pos-rate'));
    return r >= 1000 && r <= 10000 ? r : 4000;
  });
  const [menu, setMenu] = useState(() => {
    try {
      const sm = JSON.parse(localStorage.getItem('pos-menu') || 'null');
      if (Array.isArray(sm) && sm.length) return sm;
    } catch (e) {}
    return DEFAULT_MENU;
  });
  const [sales, setSales] = useState<SaleRecord[]>(() => {
    try {
      const s = JSON.parse(localStorage.getItem('pos-sales') || '[]');
      if (Array.isArray(s)) return s;
    } catch (e) {}
    return [];
  });
  const [guestOrders, setGuestOrders] = useState<GuestOrder[]>(() => {
    try {
      const g = JSON.parse(localStorage.getItem('pos-gorders') || '[]');
      if (Array.isArray(g)) return g;
    } catch (e) {}
    return [];
  });
  const [promoMsg, setPromoMsg] = useState<string>(() => localStorage.getItem('pos-promo') || "Today's specials");
  const [promoLink, setPromoLink] = useState<string>(() => localStorage.getItem('pos-link') || '');
  const [adImg, setAdImg] = useState<string>(() => localStorage.getItem('pos-ad') || '');
  const [qr, setQr] = useState<{ img: string; name: string }>(() => ({
    img: localStorage.getItem('pos-qr') || '',
    name: localStorage.getItem('pos-qrname') || ''
  }));
  const [tg, setTg] = useState<{ token: string; chat: string; auto: boolean; orderAlert: boolean }>(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('pos-tg') || '{}');
      return {
        token: saved.token || '',
        chat: saved.chat || '',
        auto: !!saved.auto,
        orderAlert: saved.orderAlert !== undefined ? !!saved.orderAlert : true
      };
    } catch (e) {
      return { token: '', chat: '', auto: false, orderAlert: true };
    }
  });
  const [pins, setPins] = useState<{ A: string; S: string }>(() => ({
    A: localStorage.getItem('pos-pinA') || '',
    S: localStorage.getItem('pos-pinS') || ''
  }));

  const [cloudConnected, setCloudConnected] = useState<boolean>(false);
  const [dbUploadMsg, setDbUploadMsg] = useState<string>('');

  // Toast
  const [toastMsg, setToastMsg] = useState<string>('');
  const toastTimerRef = useRef<any>(null);
  const showToast = (txt: string | undefined | null) => {
    if (!txt) return;
    setToastMsg(txt);
    clearTimeout(toastTimerRef.current);
    toastTimerRef.current = setTimeout(() => setToastMsg(''), 6000);
  };

  // Translation Helper
  const t = (str: string | undefined | null): string => {
    if (!str) return '';
    if (lang !== 'kh') return str;
    const trimmed = str.trim();
    if (!trimmed) return str;
    if (KH_DICT[trimmed] !== undefined) {
      return str.replace(trimmed, KH_DICT[trimmed]);
    }
    for (const [re, fn] of KH_RX) {
      const mm = re.exec(trimmed);
      if (mm) {
        return str.replace(trimmed, fn(mm));
      }
    }
    return str;
  };

  // Currency helpers
  const m = (n: number) => '$' + n.toFixed(2);
  const khr = (n: number, r: number = rate) => '៛' + (Math.round(n * r / 100) * 100).toLocaleString('en-US');
  const kh = (n: number) => khr(n, rate);
  const both = (n: number) => m(n) + ' / ' + kh(n);

  // Sync role to body dataset
  useEffect(() => {
    document.body.dataset.role = role;
    if (role) {
      sessionStorage.setItem('pos-role', role);
    } else {
      sessionStorage.removeItem('pos-role');
    }
  }, [role]);

  // Show promo pop-up when loaded as guest
  useEffect(() => {
    if (role === 'guest') {
      setPromoDlgOpen(true);
    }
  }, []);

  // Sync lang to document html
  useEffect(() => {
    document.documentElement.lang = lang === 'kh' ? 'km' : 'en';
    localStorage.setItem('pos-lang', lang);
  }, [lang]);

  // ===================== FIREBASE REAL-TIME SYNC =====================
  useEffect(() => {
    testConnection();

    // 1. Listen for Menu Items from Firestore
    const unsubItems = onSnapshot(collection(db, 'items'), (snap) => {
      setCloudConnected(true);
      if (!snap.empty) {
        const catMap: Record<string, [string, number, string, boolean, string][]> = {};
        snap.docs.forEach(docSnap => {
          const x = docSnap.data();
          if (!x || !x.cat || !x.name) return;
          if (!catMap[x.cat]) catMap[x.cat] = [];
          catMap[x.cat].push([x.name, x.price, x.img || '🍽️', !!x.promo, x.desc || '']);
        });
        const loadedMenu = Object.entries(catMap).map(([c, i]) => ({ c, i }));
        if (loadedMenu.length) {
          setMenu(loadedMenu);
          try { localStorage.setItem('pos-menu', JSON.stringify(loadedMenu)); } catch (e) {}
        }
      }
    }, (err) => {
      handleFirestoreError(err, OperationType.GET, 'items');
    });

    // 2. Listen for Sales Records from Firestore
    const unsubSales = onSnapshot(query(collection(db, 'sales'), orderBy('ts', 'desc'), limit(1000)), (snap) => {
      setCloudConnected(true);
      if (!snap.empty) {
        const loadedSales = snap.docs.map(docSnap => docSnap.data() as SaleRecord);
        setSales(loadedSales);
        try { localStorage.setItem('pos-sales', JSON.stringify(loadedSales)); } catch (e) {}
      }
    }, (err) => {
      handleFirestoreError(err, OperationType.GET, 'sales');
    });

    // 3. Listen for Guest Orders from Firestore
    const unsubGuest = onSnapshot(query(collection(db, 'guestorders'), orderBy('ts', 'desc'), limit(100)), (snap) => {
      setCloudConnected(true);
      if (!snap.empty) {
        const loadedG = snap.docs.map(docSnap => docSnap.data() as GuestOrder);
        setGuestOrders(loadedG);
        try { localStorage.setItem('pos-gorders', JSON.stringify(loadedG)); } catch (e) {}
      }
    }, (err) => {
      handleFirestoreError(err, OperationType.GET, 'guestorders');
    });

    // 4. Listen for POS Settings, Ad, and KHQR
    const unsubPos = onSnapshot(collection(db, 'pos'), (snap) => {
      setCloudConnected(true);
      snap.docs.forEach(d => {
        const x = d.data();
        if (d.id === 'settings') {
          if (typeof x.name === 'string') setRname(x.name);
          if (typeof x.promoMsg === 'string') setPromoMsg(x.promoMsg);
          if (typeof x.promoLink === 'string') setPromoLink(x.promoLink);
          if (typeof x.rate === 'number' && x.rate > 0) setRate(x.rate);
          if (typeof x.qrName === 'string') setQr(prev => ({ ...prev, name: x.qrName }));
          if (typeof x.pinA === 'string' || typeof x.pinS === 'string') {
            setPins(prev => ({ ...prev, A: x.pinA || prev.A, S: x.pinS || prev.S }));
          }
        }
        if (d.id === 'ad') {
          setAdImg(x.img || '');
        }
        if (d.id === 'khqr') {
          setQr(prev => ({ ...prev, img: x.img || '' }));
        }
      });
    }, (err) => {
      handleFirestoreError(err, OperationType.GET, 'pos');
    });

    return () => {
      unsubItems();
      unsubSales();
      unsubGuest();
      unsubPos();
    };
  }, []);

  // POS State
  const [table, setTable] = useState<string>('1');
  const [cat, setCat] = useState<number>(0);
  const [promoTab, setPromoTab] = useState<boolean>(true);
  const [tipPct, setTipPct] = useState<number>(0);
  const [orders, setOrders] = useState<Record<string, Record<string, number>>>({});
  const [gTels, setGTels] = useState<Record<string, string>>({});
  const [gPayPref, setGPayPref] = useState<Record<string, string>>({});
  const [gOrdIds, setGOrdIds] = useState<Record<string, string[]>>({});

  // Active table order
  const activeOrder = orders[table] || {};

  const getPrice = (name: string): number => {
    for (const g of menu) {
      for (const it of g.i) {
        if (it[0] === name) return it[1];
      }
    }
    return 0;
  };

  const totals = useMemo(() => {
    let sub = 0;
    for (const [k, q] of Object.entries(activeOrder)) {
      sub += getPrice(k) * q;
    }
    const tax = sub * TAX;
    const tip = sub * tipPct / 100;
    const tot = sub + tax + tip;
    return { sub, tax, tip, tot };
  }, [activeOrder, tipPct, menu]);

  const addQty = (name: string, delta: number) => {
    setOrders(prev => {
      const cur = { ...(prev[table] || {}) };
      cur[name] = (cur[name] || 0) + delta;
      if (cur[name] <= 0) delete cur[name];
      return { ...prev, [table]: cur };
    });
  };

  const promotedList = useMemo(() => {
    const list: [string, number, string, boolean, string][] = [];
    menu.forEach((g: any) => g.i.forEach((it: any) => {
      if (it[3]) list.push(it);
    }));
    return list;
  }, [menu]);

  // Modals state
  const [payDlgOpen, setPayDlgOpen] = useState<boolean>(false);
  const [payMethod, setPayMethod] = useState<string>('Card');
  const [guestTel, setGuestTel] = useState<string>('');
  const [telErr, setTelErr] = useState<string>('');
  const [cashRecv, setCashRecv] = useState<string>('');
  const [cashRecvR, setCashRecvR] = useState<string>('');

  const [recDlgOpen, setRecDlgOpen] = useState<boolean>(false);
  const [currentReceipt, setCurrentReceipt] = useState<any>(null);

  const [promoDlgOpen, setPromoDlgOpen] = useState<boolean>(false);
  const [promoPageOpen, setPromoPageOpen] = useState<boolean>(false);
  const [menuDlgOpen, setMenuDlgOpen] = useState<boolean>(false);
  const [dbDlgOpen, setDbDlgOpen] = useState<boolean>(false);
  const [tgDlgOpen, setTgDlgOpen] = useState<boolean>(false);
  const [dashOpen, setDashOpen] = useState<boolean>(false);
  const [goDlgOpen, setGoDlgOpen] = useState<boolean>(false);
  const [accDlgOpen, setAccDlgOpen] = useState<boolean>(false);
  const [pinDlgOpen, setPinDlgOpen] = useState<boolean>(false);
  const [pinTarget, setPinTarget] = useState<string>('');
  const [pinInput, setPinInput] = useState<string>('');
  const [pinErr, setPinErr] = useState<string>('');
  const [proofImg, setProofImg] = useState<string>('');

  // Audio oscillator
  const playBeep = (ok: boolean) => {
    try {
      const AudioCtx = window.AudioContext || (window as any).webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = new AudioCtx();
      const curT = ctx.currentTime;
      (ok ? [[660, 0], [880, 0.18]] : [[330, 0], [260, 0.18]]).forEach(([f, d]) => {
        const o = ctx.createOscillator();
        const gainNode = ctx.createGain();
        o.frequency.value = f;
        gainNode.gain.setValueAtTime(0.2, curT + d);
        gainNode.gain.exponentialRampToValueAtTime(0.001, curT + d + 0.25);
        o.connect(gainNode);
        gainNode.connect(ctx.destination);
        o.start(curT + d);
        o.stop(curT + d + 0.3);
      });
    } catch (e) {}
  };

  // Cash Calculation
  const recvTotal = (parseFloat(cashRecv) || 0) + (parseFloat(cashRecvR) || 0) / rate;
  const dueR = Math.round(totals.tot * rate / 100) * 100;
  let chgText = '';
  if (payMethod === 'Cash' && (cashRecv || cashRecvR)) {
    if (Math.round(recvTotal * rate) >= dueR - 50) {
      chgText = 'Change: ' + both(Math.max(0, recvTotal - totals.tot));
    } else {
      chgText = 'Short by ' + both(Math.max(0, totals.tot - recvTotal));
    }
  }

  // Telegram messaging
  const tgSend = async (text: string) => {
    if (!tg.token || !tg.chat) {
      return { ok: false, error: 'Telegram is not set up. Open the Telegram button and add the bot token and chat ID.' };
    }
    try {
      const res = await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/sendMessage`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ chat_id: tg.chat, text })
      });
      const j = await res.json();
      return j.ok ? { ok: true } : { ok: false, error: 'Telegram said: ' + (j.description || 'request failed') };
    } catch (e: any) {
      return { ok: false, error: 'Could not reach Telegram. Check your connection.' };
    }
  };

  // Function to build Telegram alert text precisely matching user's image
  const buildTelegramAlertText = (o: GuestOrder) => {
    const priorSales = o.tel
      ? sales.filter(s => s.tel && s.tel.replace(/\D/g, '') === o.tel?.replace(/\D/g, ''))
      : [];
    const priorCount = priorSales.length;
    const priorSpent = priorSales.reduce((acc, s) => acc + (s.total || 0), 0);

    const isCod = o.pay === 'cod';
    const payTitle = isCod ? 'COD' : 'KHQR';
    const payMethodStr = isCod ? 'Cash on Delivery (COD)' : 'KHQR (Scan & Pay)';

    const dateFormatted = formatTgDate(new Date(o.ts || Date.now()));
    const orderIdClean = (o.code || (o.id ? o.id.replace(/^g/, '').slice(0, 8) : '758e0885')).toLowerCase();

    const tgUserStr = o.tgUser
      ? (o.tgUser.startsWith('@') ? o.tgUser : '@' + o.tgUser)
      : '@Guest';

    const phoneStr = o.tel || '070202020';

    const lines = [
      `🔔 NEW ${payTitle} ORDER PENDING APPROVAL`,
      `---------------------------------`,
      `🆔 Order ID: ${orderIdClean}`,
      `📅 Date: ${dateFormatted}`,
      `👤 Telegram User: ${tgUserStr}`,
      `📱 Customer Phone: ${phoneStr}`,
      `📊 Customer History: ${priorCount} prior order(s) | Spent: $${priorSpent.toFixed(2)}`,
      `💳 Payment Method: ${payMethodStr}`,
      `💰 Grand Total: $${o.total.toFixed(2)}`,
      `---------------------------------`,
      `Ordered Items:`
    ];

    o.items.forEach((it, idx) => {
      const name = it[0];
      const qty = it[1];
      const priceVal = it[2] || getPrice(name);
      const noteStr = o.note ? ` (${o.note})` : '';
      lines.push(`${idx + 1}. ${name} x ${qty}${noteStr} = $${(priceVal * qty).toFixed(2)}`);
    });

    return lines.join('\n');
  };

  // Function to send Telegram order alert with inline approval/discount buttons
  const sendTelegramOrderAlert = async (o: GuestOrder) => {
    if (!tg.token || !tg.chat || tg.orderAlert === false) return { ok: false, error: 'Telegram alert not configured' };

    const text = buildTelegramAlertText(o);

    const replyMarkup = {
      inline_keyboard: [
        [{ text: "✅ Approve Order", callback_data: `appr_${o.id}` }],
        [{ text: "🏷️ Apply $1.00 Discount", callback_data: `disc1_${o.id}` }],
        [{ text: "🏷️ Apply 10% Discount", callback_data: `disc10_${o.id}` }],
        [{ text: "❌ Reject Order", callback_data: `rej_${o.id}` }]
      ]
    };

    try {
      if (o.proof && o.proof.startsWith('data:image')) {
        const blob = dataURItoBlob(o.proof);
        const fd = new FormData();
        fd.append('chat_id', tg.chat);
        fd.append('photo', blob, 'khqr_proof.jpg');
        fd.append('caption', text);
        fd.append('reply_markup', JSON.stringify(replyMarkup));

        const res = await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/sendPhoto`, {
          method: 'POST',
          body: fd
        });
        const data = await res.json();
        if (data.ok) {
          if (data.result?.message_id) {
            o.tgMsgId = data.result.message_id;
          }
          return { ok: true };
        }
      }

      const res = await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/sendMessage`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          chat_id: tg.chat,
          text,
          reply_markup: replyMarkup
        })
      });
      const data = await res.json();
      if (data.ok && data.result?.message_id) {
        o.tgMsgId = data.result.message_id;
      }
      return data.ok ? { ok: true } : { ok: false, error: data.description };
    } catch (err: any) {
      return { ok: false, error: err?.message || 'Could not reach Telegram' };
    }
  };

  // Telegram updates polling and callback query handling
  const lastUpdateIdRef = useRef<number>(0);
  const pollingActiveRef = useRef<boolean>(false);

  const freeTable = useMemo(() => {
    return TABLES.find(tName => tName !== "Takeout" && !Object.keys(orders[tName] || {}).length) || "1";
  }, [orders]);

  const handleTelegramCallbackQuery = async (cq: any) => {
    const data = String(cq.data || '');
    const fromUser = cq.from?.username ? `@${cq.from.username}` : (cq.from?.first_name || 'Admin');
    const msgId = cq.message?.message_id;
    const chatId = cq.message?.chat?.id || tg.chat;

    const answerCq = async (textMsg: string) => {
      try {
        await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/answerCallbackQuery`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ callback_query_id: cq.id, text: textMsg })
        });
      } catch (e) {}
    };

    // APPROVE ORDER
    if (data.startsWith('appr_')) {
      const orderId = data.replace('appr_', '');
      const o = guestOrders.find(item => item.id === orderId);
      if (!o) {
        await answerCq("Order not found or already processed.");
        return;
      }

      const assigned = o.table && TABLES.includes(o.table) ? o.table : freeTable;

      setOrders(prev => {
        const cur = { ...(prev[assigned] || {}) };
        o.items.forEach(([n, q]) => {
          cur[n] = (cur[n] || 0) + (Number(q) || 0);
        });
        return { ...prev, [assigned]: cur };
      });

      if (o.tel && /^\d{9,10}$/.test(o.tel)) {
        setGTels(prev => ({ ...prev, [assigned]: o.tel! }));
      }
      if (o.pay === 'khqr') {
        setGPayPref(prev => ({ ...prev, [assigned]: 'KHQR' }));
      }
      setGOrdIds(prev => ({ ...prev, [assigned]: [...(prev[assigned] || []), o.id] }));

      const updatedFields = {
        status: 'confirmed' as const,
        table: assigned,
        reason: `Approved via Telegram by ${fromUser}`,
        decidedAt: Date.now()
      };

      const updated = guestOrders.map(item => item.id === orderId ? { ...item, ...updatedFields } : item);
      setGuestOrders(updated);
      try {
        localStorage.setItem('pos-gorders', JSON.stringify(updated));
        await updateDoc(doc(db, 'guestorders', orderId), updatedFields);
      } catch (e) {}

      playBeep(true);
      showToast(`Order #${o.code} approved via Telegram (${fromUser}) -> Table ${assigned}`);
      await answerCq(`✅ Order #${o.code} approved! Table ${assigned}`);

      try {
        await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/editMessageReplyMarkup`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            chat_id: chatId,
            message_id: msgId,
            reply_markup: {
              inline_keyboard: [
                [{ text: `✅ Approved by ${fromUser} (Table ${assigned})`, callback_data: 'done' }]
              ]
            }
          })
        });
      } catch (e) {}
    }

    // APPLY $1.00 DISCOUNT
    else if (data.startsWith('disc1_')) {
      const orderId = data.replace('disc1_', '');
      const o = guestOrders.find(item => item.id === orderId);
      if (!o) {
        await answerCq("Order not found.");
        return;
      }
      const newTot = Math.max(0, +(o.total - 1).toFixed(2));
      const updatedFields = {
        total: newTot,
        discount: (o.discount || 0) + 1
      };
      const updated = guestOrders.map(item => item.id === orderId ? { ...item, ...updatedFields } : item);
      setGuestOrders(updated);
      try {
        localStorage.setItem('pos-gorders', JSON.stringify(updated));
        await updateDoc(doc(db, 'guestorders', orderId), updatedFields);
      } catch (e) {}

      showToast(`$1.00 Discount applied to Order #${o.code}! New Total: $${newTot.toFixed(2)}`);
      await answerCq(`🏷️ Applied $1.00 Discount! New Total: $${newTot.toFixed(2)}`);

      const targetOrder = { ...o, total: newTot };
      const updatedText = buildTelegramAlertText(targetOrder) + `\n🏷️ [Applied $1.00 Discount by ${fromUser}]`;
      try {
        const methodUrl = o.proof ? 'editMessageCaption' : 'editMessageText';
        const bodyPayload: any = {
          chat_id: chatId,
          message_id: msgId,
          reply_markup: {
            inline_keyboard: [
              [{ text: "✅ Approve Order", callback_data: `appr_${o.id}` }],
              [{ text: "🏷️ Applied $1.00 Discount", callback_data: `disc1_${o.id}` }],
              [{ text: "🏷️ Apply 10% Discount", callback_data: `disc10_${o.id}` }],
              [{ text: "❌ Reject Order", callback_data: `rej_${o.id}` }]
            ]
          }
        };
        if (o.proof) {
          bodyPayload.caption = updatedText;
        } else {
          bodyPayload.text = updatedText;
        }
        await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/${methodUrl}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(bodyPayload)
        });
      } catch (e) {}
    }

    // APPLY 10% DISCOUNT
    else if (data.startsWith('disc10_')) {
      const orderId = data.replace('disc10_', '');
      const o = guestOrders.find(item => item.id === orderId);
      if (!o) {
        await answerCq("Order not found.");
        return;
      }
      const newTot = Math.max(0, +(o.total * 0.9).toFixed(2));
      const updatedFields = {
        total: newTot,
        discountPercent: (o.discountPercent || 0) + 10
      };
      const updated = guestOrders.map(item => item.id === orderId ? { ...item, ...updatedFields } : item);
      setGuestOrders(updated);
      try {
        localStorage.setItem('pos-gorders', JSON.stringify(updated));
        await updateDoc(doc(db, 'guestorders', orderId), updatedFields);
      } catch (e) {}

      showToast(`10% Discount applied to Order #${o.code}! New Total: $${newTot.toFixed(2)}`);
      await answerCq(`🏷️ Applied 10% Discount! New Total: $${newTot.toFixed(2)}`);

      const targetOrder = { ...o, total: newTot };
      const updatedText = buildTelegramAlertText(targetOrder) + `\n🏷️ [Applied 10% Discount by ${fromUser}]`;
      try {
        const methodUrl = o.proof ? 'editMessageCaption' : 'editMessageText';
        const bodyPayload: any = {
          chat_id: chatId,
          message_id: msgId,
          reply_markup: {
            inline_keyboard: [
              [{ text: "✅ Approve Order", callback_data: `appr_${o.id}` }],
              [{ text: "🏷️ Apply $1.00 Discount", callback_data: `disc1_${o.id}` }],
              [{ text: "🏷️ Applied 10% Discount", callback_data: `disc10_${o.id}` }],
              [{ text: "❌ Reject Order", callback_data: `rej_${o.id}` }]
            ]
          }
        };
        if (o.proof) {
          bodyPayload.caption = updatedText;
        } else {
          bodyPayload.text = updatedText;
        }
        await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/${methodUrl}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(bodyPayload)
        });
      } catch (e) {}
    }

    // REJECT ORDER
    else if (data.startsWith('rej_')) {
      const orderId = data.replace('rej_', '');
      const o = guestOrders.find(item => item.id === orderId);
      if (!o) {
        await answerCq("Order not found.");
        return;
      }

      const updatedFields = {
        status: 'declined' as const,
        reason: `Rejected via Telegram by ${fromUser}`,
        decidedAt: Date.now()
      };

      const updated = guestOrders.map(item => item.id === orderId ? { ...item, ...updatedFields } : item);
      setGuestOrders(updated);
      try {
        localStorage.setItem('pos-gorders', JSON.stringify(updated));
        await updateDoc(doc(db, 'guestorders', orderId), updatedFields);
      } catch (e) {}

      playBeep(false);
      showToast(`Order #${o.code} rejected via Telegram (${fromUser})`);
      await answerCq(`❌ Order #${o.code} rejected.`);

      try {
        await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/editMessageReplyMarkup`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            chat_id: chatId,
            message_id: msgId,
            reply_markup: {
              inline_keyboard: [
                [{ text: `❌ Rejected by ${fromUser}`, callback_data: 'done' }]
              ]
            }
          })
        });
      } catch (e) {}
    }
  };

  // Poll Telegram for inline button clicks
  useEffect(() => {
    if (!tg.token || !tg.chat) return;

    let isMounted = true;
    const poll = async () => {
      if (pollingActiveRef.current) return;
      pollingActiveRef.current = true;
      try {
        const res = await fetch(`https://api.telegram.org/bot${encodeURIComponent(tg.token).replace(/%3A/gi, ':')}/getUpdates?offset=${lastUpdateIdRef.current + 1}&timeout=5`);
        const data = await res.json();
        if (isMounted && data.ok && Array.isArray(data.result)) {
          for (const upd of data.result) {
            lastUpdateIdRef.current = Math.max(lastUpdateIdRef.current, upd.update_id);
            if (upd.callback_query) {
              await handleTelegramCallbackQuery(upd.callback_query);
            }
          }
        }
      } catch (e) {
        // ignore network error
      } finally {
        pollingActiveRef.current = false;
      }
    };

    const interval = setInterval(poll, 4000);
    poll();

    return () => {
      isMounted = false;
      clearInterval(interval);
    };
  }, [tg.token, tg.chat, guestOrders, orders, freeTable]);

  const getRcptText = (sold: Record<string, number>, tbl: string, tObj: any, meth: string, time: string, tel: string) => {
    const e = Object.entries(sold);
    return (rname || "Restaurant") + "\n" + (/\d/.test(tbl) ? "Table " + tbl : tbl) + " - " + time + " - " + meth + (tel ? " - Tel " + tel : "") + "\n\n" +
      e.map(([n, q]) => n + " x" + q + "  " + m(getPrice(n) * q)).join("\n") + "\n\n" +
      "Subtotal " + m(tObj.sub) + "\nTax " + m(tObj.tax) + "\nTip " + m(tObj.tip) + "\nGRAND TOTAL " + m(tObj.tot) + " / " + kh(tObj.tot);
  };

  const handleCompleteSale = async () => {
    const cleanTel = guestTel.replace(/[\s-]/g, '');
    if (!/^\d{9,10}$/.test(cleanTel)) {
      setTelErr('Enter the guest telephone: 9 or 10 digits, numbers only.');
      return;
    }
    if (payMethod === 'Cash' && Math.round(recvTotal * rate) < dueR - 50) {
      showToast('Enter an amount of at least ' + both(totals.tot));
      return;
    }

    const tm = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const dt = new Date().toLocaleDateString('en-CA');
    const sold = { ...activeOrder };
    const tbl = table;

    const saleId = 's' + Date.now() + Math.random().toString(36).slice(2, 5);
    const newSale: SaleRecord = {
      id: saleId,
      table: tbl,
      method: payMethod,
      tel: cleanTel,
      total: +totals.tot.toFixed(2),
      tip: +totals.tip.toFixed(2),
      sub: +totals.sub.toFixed(2),
      tax: +totals.tax.toFixed(2),
      rate,
      date: dt,
      ts: Date.now(),
      items: Object.entries(sold).map(([n, q]) => [n, q, +(getPrice(n) * q).toFixed(2)]),
      time: tm
    };

    const nextSales = [...sales, newSale];
    setSales(nextSales);
    try {
      localStorage.setItem('pos-sales', JSON.stringify(nextSales));
      // Save sale to Firestore
      await setDoc(doc(db, 'sales', saleId), newSale);
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, 'sales');
    }

    // Clear order
    setOrders(prev => {
      const next = { ...prev };
      delete next[tbl];
      return next;
    });
    setTipPct(0);
    setPayDlgOpen(false);

    // Prepare receipt
    const rcptObj = {
      name: rname || 'Restaurant',
      table: tbl,
      time: tm,
      method: payMethod,
      tel: cleanTel,
      items: Object.entries(sold).map(([n, q]) => [n, q, +(getPrice(n) * q).toFixed(2)]),
      sub: +totals.sub.toFixed(2),
      tax: +totals.tax.toFixed(2),
      tip: +totals.tip.toFixed(2),
      tot: +totals.tot.toFixed(2),
      rate,
      text: getRcptText(sold, tbl, totals, payMethod, tm, cleanTel)
    };
    setCurrentReceipt(rcptObj);
    setRecDlgOpen(true);

    // If guest orders linked to this table, send receipt to guest
    const guestIds = gOrdIds[tbl] || [];
    if (guestIds.length) {
      const updated = guestOrders.map(o => guestIds.includes(o.id) ? { ...o, receipt: rcptObj } : o);
      setGuestOrders(updated);
      try {
        localStorage.setItem('pos-gorders', JSON.stringify(updated));
        for (const gId of guestIds) {
          await updateDoc(doc(db, 'guestorders', gId), { receipt: rcptObj });
        }
      } catch (e) {}

      setGOrdIds(prev => {
        const next = { ...prev };
        delete next[tbl];
        return next;
      });
      showToast("Receipt sent to the guest's device");
    }

    if (tg.auto) {
      tgSend(rcptObj.text).then(res => {
        showToast(res.ok ? 'Receipt sent to Telegram' : res.error);
      });
    }
  };

  // Image shrinker
  const shrinkImage = (file: File, max: number, quality: number, cb: (dataUri: string) => void) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const k = Math.min(1, max / img.width);
        const canvas = document.createElement('canvas');
        canvas.width = Math.round(img.width * k);
        canvas.height = Math.round(img.height * k);
        const ctx = canvas.getContext('2d');
        if (ctx) {
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          cb(canvas.toDataURL('image/jpeg', quality));
        }
      };
      img.src = reader.result as string;
    };
    reader.readAsDataURL(file);
  };

  const shrinkSquare = (file: File, cb: (dataUri: string) => void) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = canvas.height = 320;
        const ctx = canvas.getContext('2d');
        if (ctx) {
          const k = Math.max(320 / img.width, 320 / img.height);
          const w = img.width * k;
          const h = img.height * k;
          ctx.drawImage(img, (320 - w) / 2, (320 - h) / 2, w, h);
          cb(canvas.toDataURL('image/jpeg', 0.75));
        }
      };
      img.src = reader.result as string;
    };
    reader.readAsDataURL(file);
  };

  // Menu editor local inputs
  const [mName, setMName] = useState('');
  const [mPrice, setMPrice] = useState('');
  const [mCat, setMCat] = useState('');
  const [mPro, setMPro] = useState(false);
  const [mDesc, setMDesc] = useState('');
  const [mMsg, setMMsg] = useState('');
  const mPicRef = useRef<HTMLInputElement>(null);
  const photoTargetRef = useRef<[number, number] | null>(null);
  const fileAnyRef = useRef<HTMLInputElement>(null);

  const handleAddMenuItem = async () => {
    const n = mName.trim();
    const p = parseFloat(mPrice);
    const c = mCat.trim() || 'Other';
    if (!n) { setMMsg('Enter an item name.'); return; }
    if (!(p > 0)) { setMMsg('Enter a price greater than 0.'); return; }
    if (menu.some((g: any) => g.i.some((it: any) => it[0].toLowerCase() === n.toLowerCase()))) {
      setMMsg('An item with that name already exists.');
      return;
    }

    const finishAdd = async (img: string) => {
      const nextMenu = menu.map((g: any) => ({ ...g, i: [...g.i] }));
      let group = nextMenu.find((g: any) => g.c.toLowerCase() === c.toLowerCase());
      if (!group) {
        group = { c, i: [] };
        nextMenu.push(group);
      }
      group.i.push([n, +p.toFixed(2), img, mPro, mDesc.trim()]);
      setMenu(nextMenu);
      try {
        localStorage.setItem('pos-menu', JSON.stringify(nextMenu));
        // Save to Firestore
        const itemId = 'm_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
        await setDoc(doc(db, 'items', itemId), {
          cat: c,
          name: n,
          price: +p.toFixed(2),
          img,
          promo: mPro,
          desc: mDesc.trim(),
          o: Date.now()
        });
      } catch (e) {
        handleFirestoreError(e, OperationType.WRITE, 'items');
      }

      setMName('');
      setMPrice('');
      setMPro(false);
      setMDesc('');
      setMMsg(n + ' added.');
    };

    const file = mPicRef.current?.files?.[0];
    if (file) {
      shrinkSquare(file, finishAdd);
    } else {
      finishAdd('🍽️');
    }
  };

  // Guest order state
  const [gCode, setGCode] = useState<string>(() => {
    const saved = sessionStorage.getItem('pos-gcode');
    if (saved) return saved;
    const bytes = new Uint8Array(4);
    crypto.getRandomValues(bytes);
    const code = Array.from(bytes).map(b => CODE_CH[b % CODE_CH.length]).join('');
    sessionStorage.setItem('pos-gcode', code);
    return code;
  });
  const [gCart, setGCart] = useState<Record<string, number>>({});
  const [gCatIdx, setGCatIdx] = useState<number>(0);
  const [gPromoTab, setGPromoTab] = useState<boolean>(true);
  const [gDlgOpen, setGDlgOpen] = useState<boolean>(false);
  const [gPayMethod, setGPayMethod] = useState<'cod' | 'khqr'>('cod');
  const [gProofData, setGProofData] = useState<string>('');
  const [gProofMsg, setGProofMsg] = useState<string>('');
  const [gTel, setGTel] = useState<string>('');
  const [gTgUser, setGTgUser] = useState<string>('');
  const [gNote, setGNote] = useState<string>('');
  const [gErr, setGErr] = useState<string>('');
  const [gDoneOrderNo, setGDoneOrderNo] = useState<string>('');
  const [gMyActiveId, setGMyActiveId] = useState<string>('');
  const [gReDlgOpen, setGReDlgOpen] = useState<boolean>(false);
  const [gReProof, setGReProof] = useState<string>('');
  const [gRcptOpen, setGRcptOpen] = useState<boolean>(false);

  const activeMyGuestOrder = guestOrders.find(o => o.id === gMyActiveId);

  const gSubtotal = useMemo(() => {
    return Object.entries(gCart).reduce((acc, [name, qty]) => acc + getPrice(name) * qty, 0);
  }, [gCart, menu]);
  const gTotalCount = useMemo(() => {
    return Object.values(gCart).reduce((a, b) => a + b, 0);
  }, [gCart]);

  const handlePlaceGuestOrder = async () => {
    const cleanTel = gTel.replace(/[\s-]/g, '');
    if (cleanTel && !/^\d{9,10}$/.test(cleanTel)) {
      setGErr('Telephone must be 9 or 10 digits.');
      return;
    }
    if (gPayMethod === 'khqr') {
      if (!qr.img) {
        setGErr('KHQR is not available yet. Please choose Cash on Delivery.');
        return;
      }
      if (!gProofData) {
        setGErr('Please upload your payment screenshot.');
        return;
      }
    }

    const items = Object.entries(gCart).filter(([n]) => getPrice(n) > 0);
    if (!items.length) return;

    const id = 'g' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
    const newOrder: GuestOrder = {
      id,
      code: gCode,
      table: '',
      pay: gPayMethod,
      payStatus: gPayMethod === 'khqr' ? 'pending' : 'cod',
      proof: gPayMethod === 'khqr' ? gProofData : '',
      items: items.map(([n, q]) => [n, q, getPrice(n)]),
      total: +(gSubtotal * (1 + TAX)).toFixed(2),
      rate,
      tel: cleanTel,
      tgUser: gTgUser.trim(),
      note: gNote.trim().slice(0, 120),
      ts: Date.now(),
      date: new Date().toLocaleDateString('en-CA'),
      time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      status: 'new'
    };

    const nextOrders = [newOrder, ...guestOrders];
    setGuestOrders(nextOrders);
    try {
      localStorage.setItem('pos-gorders', JSON.stringify(nextOrders));
      // Save order to Firestore
      await setDoc(doc(db, 'guestorders', id), newOrder);
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, 'guestorders');
    }

    // Send Telegram Alert to Admin Group!
    sendTelegramOrderAlert(newOrder).then(res => {
      if (res?.ok) {
        showToast("Order alert sent to Telegram admin group!");
      }
    });

    setGMyActiveId(id);
    setGDoneOrderNo('#' + gCode);
    setGCart({});
    setGNote('');
    setGTel('');
    setGTgUser('');
    setGProofData('');
    setGProofMsg('');
    setGErr('');
  };

  // PIN Hashing
  const hashPin = async (p: string) => {
    try {
      const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('pos:' + p));
      return Array.from(new Uint8Array(b)).map(x => x.toString(16).padStart(2, '0')).join('');
    } catch (e) {
      return 'p:' + btoa(p);
    }
  };

  // Dashboard calculations
  const [dashRange, setDashRange] = useState<'today' | 'week' | 'all'>('today');
  const [telSearch, setTelSearch] = useState<string>('');

  const filteredSales = useMemo(() => {
    const todayStr = new Date().toLocaleDateString('en-CA');
    const weekAgoStr = new Date(Date.now() - 6 * 864e5).toLocaleDateString('en-CA');
    return sales.filter(x => {
      if (dashRange === 'all') return true;
      if (dashRange === 'today') return x.date === todayStr;
      return x.date >= weekAgoStr;
    });
  }, [sales, dashRange]);

  const dashStats = useMemo(() => {
    const count = filteredSales.length;
    const rev = filteredSales.reduce((acc, s) => acc + s.total, 0);
    const revK = filteredSales.reduce((acc, s) => acc + s.total * (s.rate || rate), 0);
    const tips = filteredSales.reduce((acc, s) => acc + s.tip, 0);
    const avg = count ? rev / count : 0;
    const card = filteredSales.filter(s => s.method === 'Card').reduce((a, s) => a + s.total, 0);
    const cash = filteredSales.filter(s => s.method === 'Cash').reduce((a, s) => a + s.total, 0);
    const khqr = filteredSales.filter(s => s.method === 'KHQR').reduce((a, s) => a + s.total, 0);

    const itemMap: Record<string, [number, number]> = {};
    filteredSales.forEach(s => {
      s.items?.forEach(([n, q, a]) => {
        if (!itemMap[n]) itemMap[n] = [0, 0];
        itemMap[n][0] += q;
        itemMap[n][1] += a;
      });
    });
    const items = Object.entries(itemMap).map(([n, [q, a]]) => [n, q, +a.toFixed(2)] as [string, number, number]).sort((a, b) => b[2] - a[2]);

    const buckets: Record<string, number> = {};
    filteredSales.forEach(s => {
      const k = dashRange === 'today' ? (s.ts ? new Date(s.ts).getHours() + ':00' : '?') : s.date.slice(5);
      buckets[k] = (buckets[k] || 0) + s.total;
    });

    return { count, rev, revK, tips, avg, card, cash, khqr, items, buckets };
  }, [filteredSales, dashRange, rate]);

  const guestHistory = useMemo(() => {
    const map: Record<string, SaleRecord[]> = {};
    sales.forEach(s => {
      if (!s.tel) return;
      if (!map[s.tel]) map[s.tel] = [];
      map[s.tel].push(s);
    });
    return Object.entries(map).map(([tel, o]) => ({
      tel,
      orders: o,
      total: o.reduce((acc, x) => acc + x.total, 0),
      last: o[o.length - 1]
    })).sort((a, b) => (b.last.ts || 0) - (a.last.ts || 0));
  }, [sales]);

  const filteredGuests = useMemo(() => {
    const q = telSearch.replace(/\D/g, '');
    return guestHistory.filter(g => g.tel.includes(q));
  }, [guestHistory, telSearch]);

  // Export Excel
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const sumRows = [
      ["Sales report", ""],
      ["Restaurant", rname],
      ["Period", dashRange === 'today' ? "Today" : dashRange === 'week' ? "Last 7 days" : "All time"],
      ["Orders", dashStats.count],
      ["Revenue", +dashStats.rev.toFixed(2)],
      ["Revenue (KHR)", Math.round(dashStats.revK / 100) * 100],
      ["Average order", +dashStats.avg.toFixed(2)],
      ["Card", +dashStats.card.toFixed(2)],
      ["Cash", +dashStats.cash.toFixed(2)],
      ["KHQR", +dashStats.khqr.toFixed(2)],
      ["Tips", +dashStats.tips.toFixed(2)]
    ];
    const ordRows = [
      ["Date", "Time", "Table", "Payment", "Phone", "Subtotal", "Tax", "Tip", "Total", "Total (KHR)"],
      ...filteredSales.map(x => [
        x.date,
        x.time,
        x.table,
        x.method,
        x.tel || "",
        x.sub ?? "",
        x.tax ?? "",
        x.tip,
        x.total,
        Math.round(x.total * (x.rate || rate) / 100) * 100
      ])
    ];
    const itmRows = [
      ["Item", "Qty", "Revenue"],
      ...dashStats.items
    ];
    const gstRows = [
      ["Phone", "Orders", "Total spent", "Last visit"],
      ...guestHistory.map(g => [g.tel, g.orders.length, +g.total.toFixed(2), g.last.date + " " + g.last.time])
    ];

    [
      ["Summary", sumRows],
      ["Orders", ordRows],
      ["Items", itmRows],
      ["Guests", gstRows]
    ].forEach(([name, rows]: any) => {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = rows[0]?.map(() => ({ wch: 18 }));
      XLSX.utils.book_append_sheet(wb, ws, name);
    });

    const filename = `sales-${dashRange}-${new Date().toLocaleDateString('en-CA')}.xlsx`;
    XLSX.writeFile(wb, filename);
  };

  // Export PDF
  const handleExportPDF = () => {
    const doc = new jsPDF();
    let y = 18;
    const txt = (text: string, xPos: number, sz: number, bold: boolean = false) => {
      doc.setFontSize(sz);
      doc.setFont("helvetica", bold ? "bold" : "normal");
      doc.text(String(text), xPos, y);
    };
    const nl = (h: number) => {
      y += h;
      if (y > 280) {
        doc.addPage();
        y = 18;
      }
    };
    const row = (cols: (string | number)[], xs: number[], bold: boolean = false) => {
      cols.forEach((v, i) => txt(String(v), xs[i], 10, bold));
      nl(6);
    };

    txt(rname, 14, 18, true);
    nl(8);
    txt("Sales report - " + (dashRange === 'today' ? "Today" : dashRange === 'week' ? "Last 7 days" : "All time") + " (" + new Date().toLocaleDateString('en-CA') + ")", 14, 11);
    nl(10);

    row(["Orders", dashStats.count], [14, 70]);
    row(["Revenue", m(dashStats.rev)], [14, 70]);
    row(["Revenue (KHR)", "KHR " + Math.round(dashStats.revK).toLocaleString()], [14, 70]);
    row(["Average order", m(dashStats.avg)], [14, 70]);
    row(["Tips", m(dashStats.tips)], [14, 70]);
    nl(4);

    txt("Top items", 14, 13, true);
    nl(7);
    row(["Item", "Qty", "Revenue"], [14, 100, 140], true);
    dashStats.items.slice(0, 15).forEach(r => row([r[0], r[1], m(r[2])], [14, 100, 140]));
    nl(4);

    txt("Orders", 14, 13, true);
    nl(7);
    row(["Date", "Time", "Table", "Payment", "Phone", "Total"], [14, 36, 56, 76, 102, 155], true);
    filteredSales.slice().reverse().slice(0, 60).forEach(x => {
      row([x.date, x.time, x.table, x.method, x.tel || "-", m(x.total)], [14, 36, 56, 76, 102, 155]);
    });

    doc.save(`sales-${dashRange}-${new Date().toLocaleDateString('en-CA')}.pdf`);
  };

  // Guest order decisions state
  const [gAssignTable, setGAssignTable] = useState<Record<string, string>>({});
  const [gReasons, setGReasons] = useState<Record<string, string>>({});

  const newGuestOrders = guestOrders.filter(o => o.status === 'new');
  const recentGuestOrders = guestOrders.filter(o => o.status !== 'new').sort((a, b) => (b.decidedAt || b.ts) - (a.decidedAt || a.ts)).slice(0, 8);

  // Access & PIN state
  const [adminPinIn, setAdminPinIn] = useState('');
  const [staffPinIn, setStaffPinIn] = useState('');
  const [accMsg, setAccMsg] = useState('');

  return (
    <>
      {/* HEADER */}
      <header>
        <input
          id="rname"
          aria-label="Restaurant name"
          value={rname}
          maxLength={40}
          readOnly={role !== 'admin'}
          onChange={async e => {
            const val = e.target.value;
            setRname(val);
            localStorage.setItem('pos-name', val);
            try {
              await setDoc(doc(db, 'pos', 'settings'), { name: val, rate, promoMsg, promoLink, qrName: qr.name }, { merge: true });
            } catch (err) {}
          }}
        />
        <div>
          <button id="langBtn" onClick={() => setLang(lang === 'kh' ? 'en' : 'kh')}>
            {lang === 'kh' ? 'English' : 'ភាសាខ្មែរ'}
          </button>
          <button id="goBtn" onClick={() => setGoDlgOpen(true)}>
            {t("Guest orders")}
            {newGuestOrders.length > 0 && (
              <i className="b2" id="goCount">{newGuestOrders.length}</i>
            )}
          </button>
          <button id="accBtn" className="adm" onClick={() => {
            setAccMsg((pins.A ? "Admin PIN is set. " : "Admin PIN is not set. ") + (pins.S ? "Staff PIN is set." : "Staff PIN is not set."));
            setAccDlgOpen(true);
          }}>
            {t("Access")}
          </button>
          <button id="dbBtn" className="adm" onClick={() => setDbDlgOpen(true)}>
            {t("Database")}
          </button>
          <button id="tgBtn" className="adm" onClick={() => setTgDlgOpen(true)}>
            {t("Telegram")}
          </button>
          <button id="promoBtn" className="adm" onClick={() => setPromoDlgOpen(true)}>
            {t("Promotions")}
          </button>
          <button id="menuBtn" className="adm" onClick={() => setMenuDlgOpen(true)}>
            {t("Edit menu")}
          </button>
          <button id="rep" className="adm" onClick={() => setDashOpen(true)}>
            {t("Dashboard")}
          </button>
          <button id="theme" onClick={() => {
            const root = document.documentElement;
            const dark = root.dataset.theme ? root.dataset.theme === 'dark' : window.matchMedia('(prefers-color-scheme: dark)').matches;
            root.dataset.theme = dark ? 'light' : 'dark';
          }}>
            {t("Theme")}
          </button>
          <button id="roleBtn" onClick={() => setRole('')}>
            {t("Switch page")}
          </button>
        </div>
      </header>

      {/* MAIN POS VIEW */}
      <main>
        <section>
          {/* Tables Bar */}
          <div className="tables" id="tables" aria-label="Tables">
            {TABLES.map(tName => {
              const isSel = tName === table;
              const isBusy = Object.keys(orders[tName] || {}).length > 0;
              const displayLabel = /\d/.test(tName) ? `${t("Table")} ${tName}` : t(tName);
              return (
                <button
                  key={tName}
                  className={`${isSel ? 'sel' : ''} ${isBusy ? 'busy' : ''}`}
                  onClick={() => setTable(tName)}
                >
                  {displayLabel}
                </button>
              );
            })}
          </div>

          {/* Categories Bar */}
          <div className="cats" id="cats">
            {promotedList.length > 0 && (
              <button
                className={promoTab ? 'sel' : ''}
                onClick={() => setPromoTab(true)}
              >
                {t("★ Promotions")}
              </button>
            )}
            {menu.map((g: any, i: number) => (
              <button
                key={g.c}
                className={i === cat && !promoTab ? 'sel' : ''}
                onClick={() => {
                  setPromoTab(false);
                  setCat(i);
                }}
              >
                {t(g.c)}
              </button>
            ))}
          </div>

          {/* Items Grid */}
          <div className="grid" id="grid">
            {(promoTab ? promotedList : (menu[cat] ? menu[cat].i : [])).map((it: any) => {
              const name = it[0];
              const priceVal = it[1];
              const picVal = it[2];
              const inCartQty = activeOrder[name] || 0;
              return (
                <button
                  key={name}
                  className="item"
                  onClick={() => addQty(name, 1)}
                >
                  <div className="pic">
                    {String(picVal).startsWith('data:image') ? (
                      <img src={picVal} alt="" />
                    ) : (
                      picVal
                    )}
                  </div>
                  <b>{t(name)}</b>
                  <span>{m(priceVal)}</span>
                  {inCartQty > 0 && (
                    <i className="badge" aria-label={`${inCartQty} in cart`}>
                      {inCartQty}
                    </i>
                  )}
                  <span className="add">{t("+ Add to cart")}</span>
                </button>
              );
            })}
          </div>
        </section>

        {/* CART TICKET ASIDE */}
        <aside className="ticket" aria-live="polite">
          <h2 id="tt">
            {t(/\d/.test(table) ? `Cart - Table ${table}` : `Cart - ${table}`)}
          </h2>
          <ul className="lines" id="lines">
            {Object.keys(activeOrder).length > 0 ? (
              Object.entries(activeOrder).map(([name, qty]) => (
                <li key={name}>
                  <span>{t(name)}</span>
                  <span className="q">
                    <button
                      onClick={(e) => { e.stopPropagation(); addQty(name, -1); }}
                      aria-label={`Remove one ${name}`}
                    >−</button>
                    {qty}
                    <button
                      onClick={(e) => { e.stopPropagation(); addQty(name, 1); }}
                      aria-label={`Add one ${name}`}
                    >+</button>
                  </span>
                  <span>{m(getPrice(name) * qty)}</span>
                </li>
              ))
            ) : (
              <li className="empty">{t("No items yet. Tap Add to cart on a menu item.")}</li>
            )}
          </ul>

          <div className="row">
            <span>{t("Subtotal")}</span>
            <span id="sub">{m(totals.sub)}</span>
          </div>
          <div className="row">
            <span>{t("Tax (8%)")}</span>
            <span id="tax">{m(totals.tax)}</span>
          </div>
          <div className="row">
            <span>{t("Tip")}</span>
            <span id="tip">{m(totals.tip)}</span>
          </div>

          <div className="tips" id="tips">
            {[0, 10, 15, 20].map(p => (
              <button
                key={p}
                className={p === tipPct ? 'sel' : ''}
                onClick={() => setTipPct(p)}
              >
                {p ? `${p}%` : t("No tip")}
              </button>
            ))}
          </div>

          <div className="row total">
            <span>{t("Total")}</span>
            <span id="tot">{m(totals.tot)}</span>
          </div>
          <div className="row">
            <span>{t("Total in riel")}</span>
            <b id="totK">{kh(totals.tot)}</b>
          </div>

          <div className="pay">
            <button
              className="clr"
              id="clear"
              onClick={() => {
                if (Object.keys(activeOrder).length && window.confirm("Clear this order?")) {
                  setOrders(prev => {
                    const next = { ...prev };
                    delete next[table];
                    return next;
                  });
                }
              }}
            >
              {t("Clear order")}
            </button>
            <button
              className="go"
              id="payBtn"
              disabled={!Object.keys(activeOrder).length}
              onClick={() => {
                setGuestTel(gTels[table] || '');
                setTelErr('');
                if (gPayPref[table]) {
                  setPayMethod(gPayPref[table]);
                }
                setCashRecv('');
                setCashRecvR('');
                setPayDlgOpen(true);
              }}
            >
              {t("Take payment")}
            </button>
          </div>
        </aside>
      </main>

      {/* TAKE PAYMENT DIALOG */}
      {payDlgOpen && (
        <dialog id="payDlg" open style={{ display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40 }}>
          <h2 style={{ marginTop: 0 }}>{t("Take payment")}</h2>
          <div className="row total" style={{ border: 0, margin: 0 }}>
            <span>{t("Due")}</span>
            <span id="due">{m(totals.tot)}</span>
          </div>
          <div className="row" style={{ fontSize: 20 }}>
            <span>{t("Due in riel")}</span>
            <b id="dueK">{kh(totals.tot)}</b>
          </div>
          <div className="note" id="rateNote" style={{ marginBottom: 6 }}>
            {t(`Rate: $1 = ៛${rate.toLocaleString("en-US")}`)}
          </div>

          <label className="note" htmlFor="guestTel">{t("Guest telephone (9 or 10 digits)")}</label>
          <input
            id="guestTel"
            type="tel"
            inputMode="numeric"
            autoComplete="off"
            maxLength={14}
            placeholder="0812345678"
            value={guestTel}
            onChange={e => {
              setGuestTel(e.target.value);
              setTelErr('');
            }}
          />
          <div className="note" id="telErr" role="alert" style={{ color: 'var(--hot)', minHeight: 20 }}>
            {telErr && t(telErr)}
          </div>

          <div className="tips" id="meth">
            {['Card', 'Cash', 'KHQR'].map(meth => (
              <button
                key={meth}
                className={payMethod === meth ? 'sel' : ''}
                onClick={() => setPayMethod(meth)}
              >
                {t(meth)}
              </button>
            ))}
          </div>

          {payMethod === 'Cash' && (
            <div id="cashBox">
              <label htmlFor="recv" className="note">{t("Cash received - dollars ($)")}</label>
              <input
                id="recv"
                inputMode="decimal"
                placeholder="0.00"
                value={cashRecv}
                onChange={e => setCashRecv(e.target.value)}
              />
              <label htmlFor="recvR" className="note">{t("Cash received - riel (៛)")}</label>
              <input
                id="recvR"
                inputMode="numeric"
                placeholder="0"
                value={cashRecvR}
                onChange={e => setCashRecvR(e.target.value)}
              />
              <div className="chg" id="chg">
                {chgText && t(chgText)}
              </div>
            </div>
          )}

          <div className="pay">
            <button id="cancelPay" onClick={() => setPayDlgOpen(false)}>{t("Back")}</button>
            <button className="go" id="done" onClick={handleCompleteSale}>{t("Complete sale")}</button>
          </div>
        </dialog>
      )}

      {/* RECEIPT DIALOG */}
      {recDlgOpen && currentReceipt && (
        <dialog id="recDlg" open style={{ width: 'min(520px, 94vw)', maxHeight: '88vh', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 45 }}>
          <h2 style={{ marginTop: 0 }}>{t("Receipt")}</h2>
          <div id="recBody" style={{ maxHeight: '58vh', overflow: 'auto' }}>
            <div className="rcpt">
              <h3>{currentReceipt.name}</h3>
              <p className="note">
                {t(/\d/.test(currentReceipt.table) ? `Table ${currentReceipt.table}` : currentReceipt.table)} - {currentReceipt.time} - {t("paid by")} {t(currentReceipt.method)}
                {currentReceipt.tel ? ` - Tel ${currentReceipt.tel}` : ''}
              </p>
              <div className="r3">
                <b>{t("Item")}</b>
                <b>{t("Qty")}</b>
                <b>{t("Price")}</b>
              </div>
              {currentReceipt.items.map(([n, q, p]: any) => (
                <div className="r3" key={n}>
                  <span>{t(n)}</span>
                  <span>{q}</span>
                  <span>{m(p)}</span>
                </div>
              ))}
              <div className="row">
                <span>{t("Subtotal")}</span>
                <span>{m(currentReceipt.sub)}</span>
              </div>
              <div className="row">
                <span>{t("Tax (8%)")}</span>
                <span>{m(currentReceipt.tax)}</span>
              </div>
              <div className="row">
                <span>{t("Tip")}</span>
                <span>{m(currentReceipt.tip)}</span>
              </div>
              <div className="row grand">
                <span>{t("Grand total")}</span>
                <span>{m(currentReceipt.tot)}</span>
              </div>
              <div className="row">
                <span>{t("Grand total in riel")}</span>
                <span>{khr(currentReceipt.tot, currentReceipt.rate)}</span>
              </div>
              <p className="note" style={{ margin: '4px 0 0' }}>
                $1 = ៛{currentReceipt.rate.toLocaleString('en-US')}
              </p>
            </div>
          </div>

          <div className="pay" id="rcptShare" style={{ marginBottom: 8 }}>
            <button
              id="smsBtn"
              onClick={() => {
                if (currentReceipt.tel) {
                  window.location.href = `sms:${currentReceipt.tel}?body=${encodeURIComponent(currentReceipt.text)}`;
                }
              }}
            >
              {t("Send by SMS")}
            </button>
            <button
              id="shareBtn"
              onClick={async () => {
                try {
                  if (navigator.share) {
                    await navigator.share({ title: currentReceipt.name + " receipt", text: currentReceipt.text });
                    return;
                  }
                } catch (e) {}
                try {
                  await navigator.clipboard.writeText(currentReceipt.text);
                  showToast("Receipt copied. Paste it into a message to the guest.");
                } catch (e) {
                  showToast("Could not share or copy the receipt.");
                }
              }}
            >
              {t("Share or copy")}
            </button>
          </div>

          <div className="pay" style={{ gridTemplateColumns: '1fr' }}>
            <button className="go" id="closeRec" onClick={() => setRecDlgOpen(false)}>{t("Done")}</button>
          </div>
        </dialog>
      )}

      {/* PROMOTION BANNER POPUP DIALOG */}
      {promoDlgOpen && (
        <dialog id="promoDlg" open style={{ width: 'min(560px, 94vw)', maxHeight: '90vh', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40, overflow: 'auto' }}>
          <div id="promoGrid" className="pgrid">
            {adImg || promotedList.length > 0 ? (
              <div className="pcard">
                <button
                  className="adlink"
                  onClick={() => {
                    setPromoDlgOpen(false);
                    setPromoPageOpen(true);
                  }}
                  aria-label="Open promotion page"
                  style={{ width: '100%' }}
                >
                  {adImg ? (
                    <img className="adpic" src={adImg} alt="Promotion" />
                  ) : (
                    <div className="pic">
                      {String(promotedList[0][2]).startsWith('data:image') ? (
                        <img src={promotedList[0][2]} alt="" />
                      ) : (
                        promotedList[0][2]
                      )}
                    </div>
                  )}
                  {promotedList[0] && (
                    <div style={{ marginTop: 6 }}>
                      <b style={{ fontSize: 18 }}>{t(promotedList[0][0])}</b>
                      <span style={{ color: 'var(--mute)', marginLeft: 8 }}>{m(promotedList[0][1])}</span>
                      {promotedList[0][4] && <p className="note" style={{ margin: '4px 0 0' }}>{t(promotedList[0][4])}</p>}
                    </div>
                  )}
                </button>
              </div>
            ) : (
              <p className="empty">{t("No advertisement yet. Open Edit menu and upload an advertisement picture, or promote an item.")}</p>
            )}
          </div>
          <div className="pay" style={{ gridTemplateColumns: '1fr', marginTop: 12, gap: 8 }}>
            <button
              className="go"
              style={{ width: '100%', padding: '12px 14px', fontSize: 16 }}
              onClick={() => {
                setPromoDlgOpen(false);
                setPromoPageOpen(true);
              }}
            >
              🎉 {t("View promotion page")}
            </button>
            {promoLink && (
              <a
                href={promoLink}
                target="_blank"
                rel="noopener noreferrer"
                style={{ textAlign: 'center', fontSize: 14, color: 'var(--mute)', textDecoration: 'underline', padding: '4px 0' }}
              >
                {t("Visit promotion website ↗")}
              </a>
            )}
            <button id="closePromo" onClick={() => setPromoDlgOpen(false)}>{t("Close")}</button>
          </div>
        </dialog>
      )}

      {/* PROMOTIONS FULL PAGE VIEW */}
      {promoPageOpen && (
        <div id="promoPage" role="dialog" aria-label="Food promotions">
          <div style={{ maxWidth: 920, margin: 'auto', paddingTop: 14 }}>
            <button id="backPromo" onClick={() => setPromoPageOpen(false)}>
              {role === 'guest' ? t("Back") : t("Back to POS")}
            </button>
            <h1 id="ppHead" style={{ margin: '14px 0 4px' }}>{promoMsg}</h1>
            <p className="note" style={{ margin: '0 0 14px' }}>
              {t("Promotion categories. Add any item to the cart for the open table.")}
            </p>
            <div id="ppGrid" className="ppgrid">
              {menu.map((g: any) => {
                const pItems = g.i.filter((it: any) => it[3]);
                if (!pItems.length) return null;
                return (
                  <React.Fragment key={g.c}>
                    <h2>{t(g.c)}</h2>
                    {pItems.map((it: any) => (
                      <div className="pcard" key={it[0]}>
                        <div className="pic">
                          {String(it[2]).startsWith('data:image') ? <img src={it[2]} alt="" /> : it[2]}
                        </div>
                        <b>{t(it[0])}</b>
                        <span>{m(it[1])}</span>
                        <p className="note">{t(it[4] || '')}</p>
                        <button
                          className="add"
                          onClick={(e) => {
                            if (role === 'guest') {
                              setGCart(prev => ({ ...prev, [it[0]]: (prev[it[0]] || 0) + 1 }));
                            } else {
                              addQty(it[0], 1);
                            }
                            (e.target as HTMLElement).textContent = t("Added - tap to add another");
                          }}
                        >
                          {t("+ Add to cart")}
                        </button>
                      </div>
                    ))}
                  </React.Fragment>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* EDIT MENU DIALOG */}
      {menuDlgOpen && (
        <dialog id="menuDlg" open style={{ width: 'min(520px, 94vw)', maxHeight: '90vh', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40, overflow: 'auto' }}>
          <h2 style={{ marginTop: 0 }}>{t("Edit menu")}</h2>
          <div className="mform">
            <input
              id="mName"
              placeholder={t("Item name")}
              maxLength={40}
              value={mName}
              onChange={e => setMName(e.target.value)}
            />
            <input
              id="mPrice"
              placeholder={t("Price, e.g. 9.50")}
              inputMode="decimal"
              value={mPrice}
              onChange={e => setMPrice(e.target.value)}
            />
            <input
              id="mCat"
              list="catlist"
              placeholder={t("Category")}
              maxLength={20}
              value={mCat}
              onChange={e => setMCat(e.target.value)}
            />
            <datalist id="catlist">
              {menu.map((g: any) => <option key={g.c} value={g.c} />)}
            </datalist>
            <label className="note">
              <input
                type="checkbox"
                id="mPro"
                style={{ width: 'auto' }}
                checked={mPro}
                onChange={e => setMPro(e.target.checked)}
              /> {t("Promote this item in the pop-up")}
            </label>
            <input
              id="mDesc"
              placeholder={t("Promo description (optional)")}
              maxLength={80}
              value={mDesc}
              onChange={e => setMDesc(e.target.value)}
            />
            <label className="note" htmlFor="mPic">{t("Photo (optional)")}</label>
            <input id="mPic" type="file" accept="image/*" ref={mPicRef} />
            <button className="go" id="addItem" onClick={handleAddMenuItem}>{t("Add item")}</button>
            <div className="note" id="mMsg" role="status">{mMsg && t(mMsg)}</div>
          </div>

          <label className="note" htmlFor="mPromo">{t("Promotion headline")}</label>
          <input
            id="mPromo"
            maxLength={60}
            style={{ fontSize: 16, marginBottom: 10 }}
            value={promoMsg}
            onChange={async e => {
              const val = e.target.value;
              setPromoMsg(val);
              localStorage.setItem('pos-promo', val);
              try {
                await setDoc(doc(db, 'pos', 'settings'), { promoMsg: val }, { merge: true });
              } catch (err) {}
            }}
          />

          <label className="note" htmlFor="mRate">{t("Exchange rate (riel per 1 dollar)")}</label>
          <input
            id="mRate"
            inputMode="numeric"
            style={{ fontSize: 16, marginBottom: 10 }}
            value={rate}
            onChange={async e => {
              const val = parseInt(e.target.value, 10);
              if (val >= 1000 && val <= 10000) {
                setRate(val);
                localStorage.setItem('pos-rate', String(val));
                try {
                  await setDoc(doc(db, 'pos', 'settings'), { rate: val }, { merge: true });
                } catch (err) {}
              }
            }}
          />

          <label className="note" htmlFor="mQr">{t("KHQR payment code image (for guests to scan)")}</label>
          <input
            id="mQr"
            type="file"
            accept="image/*"
            style={{ fontSize: 14, marginBottom: 6 }}
            onChange={e => {
              const f = e.target.files?.[0];
              if (f) {
                shrinkImage(f, 640, 0.9, async data => {
                  setQr(prev => {
                    const next = { ...prev, img: data };
                    localStorage.setItem('pos-qr', data);
                    return next;
                  });
                  try {
                    await setDoc(doc(db, 'pos', 'khqr'), { img: data });
                  } catch (err) {}
                  showToast("KHQR image saved.");
                });
              }
            }}
          />
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 6 }}>
            <button
              id="qrClear"
              type="button"
              onClick={async () => {
                setQr(prev => {
                  const next = { ...prev, img: '' };
                  localStorage.removeItem('pos-qr');
                  return next;
                });
                try {
                  await deleteDoc(doc(db, 'pos', 'khqr'));
                } catch (err) {}
                showToast("KHQR image removed.");
              }}
            >
              {t("Remove KHQR image")}
            </button>
            <span className="note" id="qrMsg">
              {qr.img ? "A KHQR image is saved." : "None saved yet."}
            </span>
          </div>

          <label className="note" htmlFor="mQrName">{t("KHQR account name (shown to guests)")}</label>
          <input
            id="mQrName"
            maxLength={40}
            style={{ fontSize: 16, marginBottom: 10 }}
            value={qr.name}
            onChange={async e => {
              const v = e.target.value;
              setQr(prev => {
                const next = { ...prev, name: v };
                localStorage.setItem('pos-qrname', v);
                return next;
              });
              try {
                await setDoc(doc(db, 'pos', 'settings'), { qrName: v }, { merge: true });
              } catch (err) {}
            }}
          />

          <label className="note" htmlFor="mAd">{t("Advertisement picture for the pop-up (one picture)")}</label>
          <input
            id="mAd"
            type="file"
            accept="image/*"
            style={{ fontSize: 14, marginBottom: 6 }}
            onChange={e => {
              const f = e.target.files?.[0];
              if (f) {
                shrinkImage(f, 720, 0.8, async data => {
                  setAdImg(data);
                  localStorage.setItem('pos-ad', data);
                  try {
                    await setDoc(doc(db, 'pos', 'ad'), { img: data });
                  } catch (err) {}
                  showToast("Advertisement picture saved.");
                });
              }
            }}
          />
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
            <button
              id="adClear"
              type="button"
              onClick={async () => {
                setAdImg('');
                localStorage.removeItem('pos-ad');
                try {
                  await deleteDoc(doc(db, 'pos', 'ad'));
                } catch (err) {}
                showToast("Removed ad picture.");
              }}
            >
              {t("Remove ad picture")}
            </button>
            <span className="note" id="adMsg">
              {adImg ? "An advertisement picture is saved." : "None saved yet."}
            </span>
          </div>

          <label className="note" htmlFor="mLink">{t("Promotion page link (optional, starts with https://)")}</label>
          <input
            id="mLink"
            inputMode="url"
            placeholder={t("Leave empty to use the built-in promotions page")}
            style={{ fontSize: 16, marginBottom: 10 }}
            value={promoLink}
            onChange={async e => {
              const val = e.target.value;
              setPromoLink(val);
              localStorage.setItem('pos-link', val);
              try {
                await setDoc(doc(db, 'pos', 'settings'), { promoLink: val }, { merge: true });
              } catch (err) {}
            }}
          />

          {/* Existing menu items list */}
          <div id="mList" style={{ maxHeight: '38vh', overflow: 'auto' }}>
            {menu.map((g: any, gi: number) => (
              <React.Fragment key={g.c}>
                <div className="mrow" style={{ gridTemplateColumns: '1fr auto' }}>
                  <b>{t(g.c)}</b>
                  <button
                    onClick={() => {
                      const next = menu.map((grp: any, idx: number) => {
                        if (idx !== gi) return grp;
                        const allPromoted = grp.i.every((it: any) => it[3]);
                        return { ...grp, i: grp.i.map((it: any) => [it[0], it[1], it[2], !allPromoted, it[4]]) };
                      });
                      setMenu(next);
                      localStorage.setItem('pos-menu', JSON.stringify(next));
                    }}
                  >
                    {g.i.every((it: any) => it[3]) ? t("Unpromote category") : t("Promote category")}
                  </button>
                </div>
                {g.i.map((it: any, ii: number) => (
                  <div className="mrow" key={it[0]}>
                    <div className="pic">
                      {String(it[2]).startsWith('data:image') ? <img src={it[2]} alt="" /> : it[2]}
                    </div>
                    <span>
                      {t(it[0])}<br />
                      <span className="note">{t(g.c)} - {m(it[1])}</span>
                    </span>
                    <button
                      onClick={() => {
                        const next = menu.map((grp: any, gIdx: number) => {
                          if (gIdx !== gi) return grp;
                          return {
                            ...grp,
                            i: grp.i.map((item: any, iIdx: number) => iIdx === ii ? [item[0], item[1], item[2], !item[3], item[4]] : item)
                          };
                        });
                        setMenu(next);
                        localStorage.setItem('pos-menu', JSON.stringify(next));
                      }}
                    >
                      {it[3] ? t("★ Promoted") : t("Promote")}
                    </button>
                    <button
                      onClick={() => {
                        photoTargetRef.current = [gi, ii];
                        fileAnyRef.current?.click();
                      }}
                    >
                      {t("Photo")}
                    </button>
                    <button
                      style={{ color: 'var(--hot)' }}
                      onClick={async () => {
                        if (window.confirm("Remove " + it[0] + " from the menu?")) {
                          const next = menu.map((grp: any, gIdx: number) => {
                            if (gIdx !== gi) return grp;
                            return { ...grp, i: grp.i.filter((_: any, iIdx: number) => iIdx !== ii) };
                          }).filter((grp: any) => grp.i.length > 0);
                          setMenu(next);
                          localStorage.setItem('pos-menu', JSON.stringify(next));

                          // Delete from Firestore
                          try {
                            const snap = await getDocs(collection(db, 'items'));
                            snap.docs.forEach(d => {
                              if (d.data().name === it[0]) {
                                deleteDoc(doc(db, 'items', d.id));
                              }
                            });
                          } catch (err) {}
                        }
                      }}
                    >
                      {t("Remove")}
                    </button>
                  </div>
                ))}
              </React.Fragment>
            ))}
          </div>

          <input
            id="fileAny"
            type="file"
            accept="image/*"
            hidden
            ref={fileAnyRef}
            onChange={e => {
              const file = e.target.files?.[0];
              if (file && photoTargetRef.current) {
                const [gi, ii] = photoTargetRef.current;
                shrinkSquare(file, data => {
                  const next = menu.map((grp: any, gIdx: number) => {
                    if (gIdx !== gi) return grp;
                    return {
                      ...grp,
                      i: grp.i.map((item: any, iIdx: number) => iIdx === ii ? [item[0], item[1], data, item[3], item[4]] : item)
                    };
                  });
                  setMenu(next);
                  localStorage.setItem('pos-menu', JSON.stringify(next));
                });
              }
            }}
          />

          <div className="pay" style={{ gridTemplateColumns: '1fr', marginTop: 10 }}>
            <button id="closeMenu" onClick={() => setMenuDlgOpen(false)}>{t("Done")}</button>
          </div>
        </dialog>
      )}

      {/* SALES DASHBOARD VIEW */}
      {dashOpen && (
        <div id="dash" role="dialog" aria-label="Sales dashboard">
          <div style={{ maxWidth: 980, margin: 'auto', paddingTop: 14 }}>
            <div className="dbar">
              <button id="backDash" onClick={() => setDashOpen(false)}>{t("Back to POS")}</button>
              <span style={{ flex: 1 }} />
              <button id="xlsBtn" onClick={handleExportExcel}>{t("Export Excel")}</button>
              <button id="pdfBtn" onClick={handleExportPDF}>{t("Export PDF")}</button>
              <button
                id="dTg"
                onClick={async () => {
                  if (!filteredSales.length) {
                    showToast("No sales in this range to send.");
                    return;
                  }
                  const txt = `${rname} - sales summary (${dashRange})\nOrders: ${dashStats.count}\nRevenue: ${m(dashStats.rev)} / ៛${Math.round(dashStats.revK).toLocaleString()}\nCard: ${m(dashStats.card)}\nCash: ${m(dashStats.cash)}\nKHQR: ${m(dashStats.khqr)}\nTips: ${m(dashStats.tips)}`;
                  const res = await tgSend(txt);
                  showToast(res.ok ? "Summary sent to Telegram" : res.error);
                }}
              >
                {t("Send to Telegram")}
              </button>
            </div>

            <h1 style={{ margin: '14px 0 8px' }}>{t("Sales dashboard")}</h1>
            <div className="tips" id="dRange" style={{ maxWidth: 440 }}>
              {(['today', 'week', 'all'] as const).map(rng => (
                <button
                  key={rng}
                  className={rng === dashRange ? 'sel' : ''}
                  onClick={() => setDashRange(rng)}
                >
                  {t(rng === 'today' ? "Today" : rng === 'week' ? "Last 7 days" : "All time")}
                </button>
              ))}
            </div>

            <div className="kpis" id="kpis">
              <div className="kpi">
                <b>{m(dashStats.rev)}</b>
                <span>{t("Revenue")}</span>
              </div>
              <div className="kpi">
                <b>{"៛" + (Math.round(dashStats.revK / 100) * 100).toLocaleString('en-US')}</b>
                <span>{t("Revenue (KHR)")}</span>
              </div>
              <div className="kpi">
                <b>{dashStats.count}</b>
                <span>{t("Orders")}</span>
              </div>
              <div className="kpi">
                <b>{m(dashStats.avg)}</b>
                <span>{t("Average order")}</span>
              </div>
              <div className="kpi">
                <b>{m(dashStats.tips)}</b>
                <span>{t("Tips")}</span>
              </div>
            </div>

            <div className="dgrid">
              {/* Revenue Chart */}
              <section className="card">
                <h2 id="revT">
                  {t(dashRange === 'today' ? "Revenue by hour" : "Revenue by day")}
                </h2>
                <div id="chRev">
                  {Object.keys(dashStats.buckets).length > 0 ? (
                    (() => {
                      const keys = Object.keys(dashStats.buckets).sort((a, b) => dashRange === 'today' ? parseInt(a) - parseInt(b) : a.localeCompare(b));
                      const vals = keys.map(k => dashStats.buckets[k]);
                      const mx = Math.max(...vals, 1);
                      const W = 560;
                      const H = 190;
                      const pad = 16;
                      const step = (W - pad * 2) / vals.length;
                      const bw = Math.min(44, step - 6);
                      return (
                        <svg viewBox={`0 0 ${W} ${H}`} width="100%" role="img" aria-label="Revenue chart">
                          {vals.map((v, i) => {
                            const xPos = pad + i * step + (step - bw) / 2;
                            const h = (H - 56) * v / mx;
                            return (
                              <g key={keys[i]}>
                                <rect x={xPos} y={H - 30 - h} width={bw} height={h} rx={4} fill="var(--acc)" />
                                <text x={xPos + bw / 2} y={H - 14} textAnchor="middle" fontSize={11} fill="var(--mute)">
                                  {keys[i]}
                                </text>
                                <text x={xPos + bw / 2} y={H - 34 - h} textAnchor="middle" fontSize={11} fill="var(--ink)">
                                  {Math.round(v)}
                                </text>
                              </g>
                            );
                          })}
                        </svg>
                      );
                    })()
                  ) : (
                    <p className="empty">{t("No sales in this range yet.")}</p>
                  )}
                </div>
              </section>

              {/* Top Items */}
              <section className="card">
                <h2>{t("Top items")}</h2>
                <div id="chTop">
                  {dashStats.items.length > 0 ? (
                    dashStats.items.slice(0, 8).map(r => {
                      const maxQty = dashStats.items[0][1] || 1;
                      return (
                        <div className="hb" key={r[0]}>
                          <span>{t(r[0])}</span>
                          <i style={{ width: `${Math.max(3, (r[1] / maxQty) * 100)}%` }} />
                          <span>{r[1]} {t("sold")}</span>
                        </div>
                      );
                    })
                  ) : (
                    <p className="empty">{t("No data yet.")}</p>
                  )}
                </div>
              </section>

              {/* Payment Methods */}
              <section className="card">
                <h2>{t("Payment methods")}</h2>
                <div id="chPay">
                  {[
                    [t("Card"), dashStats.card, m(dashStats.card)],
                    [t("Cash"), dashStats.cash, m(dashStats.cash)],
                    [t("KHQR"), dashStats.khqr, m(dashStats.khqr)]
                  ].map(([label, val, display]: any) => {
                    const maxP = Math.max(dashStats.card, dashStats.cash, dashStats.khqr, 1);
                    return (
                      <div className="hb" key={label}>
                        <span>{label}</span>
                        <i style={{ width: `${Math.max(3, (val / maxP) * 100)}%` }} />
                        <span>{display}</span>
                      </div>
                    );
                  })}
                </div>
              </section>

              {/* Recent Orders */}
              <section className="card">
                <h2>{t("Recent orders")}</h2>
                <div id="chOrd" style={{ overflowX: 'auto' }}>
                  {filteredSales.length > 0 ? (
                    <table>
                      <thead>
                        <tr>
                          <th>{t("Date")}</th>
                          <th>{t("Time")}</th>
                          <th>{t("Table")}</th>
                          <th>{t("Pay")}</th>
                          <th>{t("Phone")}</th>
                          <th>{t("Total")}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredSales.slice().reverse().slice(0, 10).map((x, idx) => (
                          <tr key={idx}>
                            <td>{x.date.slice(5)}</td>
                            <td>{x.time}</td>
                            <td>{t(/\d/.test(x.table) ? `Table ${x.table}` : x.table)}</td>
                            <td>{t(x.method)}</td>
                            <td>{x.tel || "-"}</td>
                            <td>{m(x.total)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <p className="empty">{t("No orders in this range yet.")}</p>
                  )}
                </div>
              </section>

              {/* Guest orders history by telephone */}
              <section className="card" style={{ gridColumn: '1 / -1' }}>
                <h2>{t("Guest order history by telephone")}</h2>
                <p className="note" style={{ margin: '0 0 8px' }}>
                  {t("All saved orders, newest guest first. Tap a number to expand.")}
                </p>
                <input
                  id="telSearch"
                  inputMode="numeric"
                  placeholder={t("Search telephone")}
                  aria-label="Search telephone"
                  style={{ fontSize: 16, padding: '8px 10px', border: '1px solid var(--line)', borderRadius: 10, background: 'var(--bg)', color: 'var(--ink)', width: '100%', maxWidth: 300, marginBottom: 8 }}
                  value={telSearch}
                  onChange={e => setTelSearch(e.target.value)}
                />
                <div id="chGuests">
                  {filteredGuests.length > 0 ? (
                    filteredGuests.map(g => (
                      <details className="gst" key={g.tel}>
                        <summary>
                          <b>{g.tel}</b>
                          <span>
                            {g.orders.length} {t(g.orders.length > 1 ? "orders" : "order")} - {m(g.total)} - {t("last visit")} {g.last.date}
                          </span>
                        </summary>
                        <div style={{ overflowX: 'auto' }}>
                          <table>
                            <thead>
                              <tr>
                                <th>{t("Date")}</th>
                                <th>{t("Time")}</th>
                                <th>{t("Table")}</th>
                                <th>{t("Pay")}</th>
                                <th>{t("Items")}</th>
                                <th>{t("Total")}</th>
                              </tr>
                            </thead>
                            <tbody>
                              {g.orders.slice().reverse().map((o, idx) => (
                                <tr key={idx}>
                                  <td>{o.date}</td>
                                  <td>{o.time}</td>
                                  <td>{t(/\d/.test(o.table) ? `Table ${o.table}` : o.table)}</td>
                                  <td>{t(o.method)}</td>
                                  <td>{o.items?.map(it => `${t(it[0])} x${it[1]}`).join(', ') || '-'}</td>
                                  <td>{m(o.total)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </details>
                    ))
                  ) : (
                    <p className="empty">
                      {telSearch ? t("No guest matches that number.") : t("No orders with a telephone number yet.")}
                    </p>
                  )}
                </div>
              </section>
            </div>

            <button
              id="clearSales"
              style={{ marginTop: 14, color: 'var(--hot)', borderColor: 'var(--hot)' }}
              onClick={async () => {
                if (window.confirm("Delete ALL saved sales data? This cannot be undone.")) {
                  setSales([]);
                  localStorage.removeItem('pos-sales');
                  try {
                    const snap = await getDocs(collection(db, 'sales'));
                    snap.docs.forEach(d => deleteDoc(doc(db, 'sales', d.id)));
                  } catch (e) {}
                  showToast("Cleared all sales records.");
                }
              }}
            >
              {t("Clear all sales data")}
            </button>
          </div>
        </div>
      )}

      {/* GATE SCREEN (Choose Role) */}
      {!role && (
        <div id="gate" role="dialog" aria-label="Choose a page">
          <div className="gbox">
            <h1 style={{ margin: '0 0 14px' }}>{t("Choose a page")}</h1>
            <div className="gcards">
              <button
                data-role="guest"
                onClick={() => {
                  setRole('guest');
                  setPromoDlgOpen(true);
                }}
              >
                <b>{t("Guest order")}</b>
                <span>{t("Order from the table")}</span>
              </button>
              <button
                data-role="staff"
                onClick={() => {
                  if (!pins.S && !pins.A) {
                    setRole('staff');
                  } else {
                    setPinTarget('staff');
                    setPinInput('');
                    setPinErr('');
                    setPinDlgOpen(true);
                  }
                }}
              >
                <b>{t("Staff")}</b>
                <span>{t("Register and guest orders")}</span>
              </button>
              <button
                data-role="admin"
                onClick={() => {
                  if (!pins.A) {
                    setRole('admin');
                  } else {
                    setPinTarget('admin');
                    setPinInput('');
                    setPinErr('');
                    setPinDlgOpen(true);
                  }
                }}
              >
                <b>{t("Admin")}</b>
                <span>{t("Menu, reports and settings")}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* GUEST ORDERING PAGE */}
      {role === 'guest' && (
        <div id="guestPage">
          <div className="gwrap">
            <div className="gtop">
              <b id="gName">{rname || "Restaurant"}</b>
              <span style={{ flex: 1 }} />
              <button id="gPromoTop" onClick={() => setPromoPageOpen(true)}>
                {t("★ Promotions")}
              </button>
              <button id="gLang" onClick={() => setLang(lang === 'kh' ? 'en' : 'kh')}>
                {lang === 'kh' ? 'English' : 'ភាសាខ្មែរ'}
              </button>
              <button id="gStaff" onClick={() => setRole('')}>
                {t("Staff login")}
              </button>
            </div>

            {/* Active order status banner */}
            {activeMyGuestOrder && (
              <>
                <div
                  id="gStatus"
                  className={`gstat ${activeMyGuestOrder.status === 'confirmed' ? 'ok' : activeMyGuestOrder.status === 'declined' ? 'bad' : 'wait'}`}
                >
                  {activeMyGuestOrder.status === 'confirmed' && (
                    <span>{t(`Order #${activeMyGuestOrder.code} confirmed.`)}{activeMyGuestOrder.table ? ` តុ ${activeMyGuestOrder.table}។` : ''}{activeMyGuestOrder.reason ? ` ${t("Note")}: ${activeMyGuestOrder.reason}` : ''}</span>
                  )}
                  {activeMyGuestOrder.status === 'declined' && (
                    <span>{t(`Order #${activeMyGuestOrder.code} was declined.`)}{activeMyGuestOrder.reason ? ` ${t("Reason")}: ${activeMyGuestOrder.reason}` : ''}</span>
                  )}
                  {activeMyGuestOrder.status === 'new' && (
                    <span>{t(`Order #${activeMyGuestOrder.code} sent. Waiting for staff to confirm.`)}</span>
                  )}
                </div>

                {activeMyGuestOrder.status !== 'declined' && (
                  <div
                    id="gPayStat"
                    className={`gstat ${activeMyGuestOrder.payStatus === 'matched' ? 'ok' : activeMyGuestOrder.payStatus === 'mismatch' ? 'bad' : 'wait'}`}
                  >
                    {activeMyGuestOrder.pay === 'khqr' ? (
                      activeMyGuestOrder.payStatus === 'matched'
                        ? t("KHQR payment matched. Thank you!")
                        : activeMyGuestOrder.payStatus === 'mismatch'
                        ? t("KHQR payment is not matching.") + (activeMyGuestOrder.payNote ? ` ${t("Reason")}: ${activeMyGuestOrder.payNote}` : '')
                        : t("KHQR payment: staff are checking your screenshot.")
                    ) : (
                      t("Pay with cash when your order is delivered.")
                    )}
                  </div>
                )}

                {activeMyGuestOrder.payStatus === 'mismatch' && (
                  <div id="gPayFix" className="pfix">
                    <button id="gReBtn" onClick={() => { setGReProof(''); setGReDlgOpen(true); }}>
                      {t("Upload new screenshot")}
                    </button>
                    <button
                      id="gCodBtn"
                      onClick={async () => {
                        const updated = guestOrders.map(o => o.id === gMyActiveId ? { ...o, pay: 'cod', payStatus: 'cod' } : o);
                        setGuestOrders(updated);
                        localStorage.setItem('pos-gorders', JSON.stringify(updated));
                        try {
                          await updateDoc(doc(db, 'guestorders', gMyActiveId), { pay: 'cod', payStatus: 'cod' });
                        } catch (e) {}
                      }}
                    >
                      {t("Pay by Cash on Delivery")}
                    </button>
                  </div>
                )}

                {activeMyGuestOrder.receipt && (
                  <button
                    id="gRcptBtn"
                    style={{ margin: '0 0 10px' }}
                    onClick={() => setGRcptOpen(true)}
                  >
                    {t("View receipt")}
                  </button>
                )}
              </>
            )}

            <p className="note" style={{ margin: '0 0 8px' }}>
              {t("Your order code")} <b id="gCode" style={{ fontSize: 22, letterSpacing: 2 }}>#{gCode}</b>
            </p>

            <div className="cats" id="gCats">
              {promotedList.length > 0 && (
                <button
                  className={gPromoTab ? 'sel' : ''}
                  onClick={() => setGPromoTab(true)}
                >
                  {t("★ Promotions")}
                </button>
              )}
              {menu.map((g: any, i: number) => (
                <button
                  key={g.c}
                  className={i === gCatIdx && !gPromoTab ? 'sel' : ''}
                  onClick={() => {
                    setGPromoTab(false);
                    setGCatIdx(i);
                  }}
                >
                  {t(g.c)}
                </button>
              ))}
            </div>

            <div className="grid" id="gGrid">
              {(gPromoTab ? promotedList : (menu[gCatIdx] ? menu[gCatIdx].i : [])).map((it: any) => {
                const name = it[0];
                const priceVal = it[1];
                const picVal = it[2];
                const qty = gCart[name] || 0;
                return (
                  <div className="gitem" key={name}>
                    <div className="pic">
                      {String(picVal).startsWith('data:image') ? <img src={picVal} alt="" /> : picVal}
                    </div>
                    <b>{t(name)}</b>
                    <span>{m(priceVal)} / {kh(priceVal)}</span>
                    <div className="gq">
                      <button
                        onClick={() => {
                          setGCart(prev => {
                            const cur = { ...prev };
                            if (cur[name] > 1) cur[name]--;
                            else delete cur[name];
                            return cur;
                          });
                        }}
                        aria-label={`Remove one ${name}`}
                      >−</button>
                      <b>{qty}</b>
                      <button
                        onClick={() => {
                          setGCart(prev => ({ ...prev, [name]: (prev[name] || 0) + 1 }));
                        }}
                        aria-label={`Add one ${name}`}
                      >+</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {gTotalCount > 0 && (
            <div className="gbar" id="gBar">
              <b id="gSum">
                {t(`${gTotalCount} items - ${both(gSubtotal * (1 + TAX))}`)}
              </b>
              <button
                className="go"
                id="gReview"
                onClick={() => {
                  setGErr('');
                  setGDlgOpen(true);
                }}
              >
                {t("View order")}
              </button>
            </div>
          )}

          {/* GUEST CART REVIEW DIALOG */}
          {gDlgOpen && (
            <dialog id="gDlg" open style={{ width: 'min(460px, 94vw)', maxHeight: '90vh', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40, overflow: 'auto' }}>
              {!gDoneOrderNo ? (
                <div id="gForm">
                  <h2 style={{ marginTop: 0 }}>{t("Your order")}</h2>
                  <ul className="lines" id="gLines">
                    {Object.entries(gCart).map(([name, qty]) => (
                      <li key={name}>
                        <span>{t(name)}</span>
                        <span className="q">
                          <button onClick={() => {
                            setGCart(prev => {
                              const cur = { ...prev };
                              if (cur[name] > 1) cur[name]--;
                              else delete cur[name];
                              return cur;
                            });
                          }}>−</button>
                          {qty}
                          <button onClick={() => {
                            setGCart(prev => ({ ...prev, [name]: (prev[name] || 0) + 1 }));
                          }}>+</button>
                        </span>
                        <span>{m(getPrice(name) * qty)}</span>
                      </li>
                    ))}
                  </ul>

                  <div id="gTotals">
                    <div className="row">
                      <span>{t("Subtotal")}</span>
                      <span>{m(gSubtotal)}</span>
                    </div>
                    <div className="row">
                      <span>{t("Tax (8%)")}</span>
                      <span>{m(gSubtotal * TAX)}</span>
                    </div>
                    <div className="row total">
                      <span>{t("Total")}</span>
                      <span>{both(gSubtotal * (1 + TAX))}</span>
                    </div>
                  </div>

                  <div className="mform" style={{ marginTop: 10 }}>
                    <b>{t("Payment")}</b>
                    <label>
                      <input
                        type="radio"
                        name="gPay"
                        value="cod"
                        checked={gPayMethod === 'cod'}
                        onChange={() => setGPayMethod('cod')}
                        style={{ width: 'auto' }}
                      /> {t("Cash on Delivery")}
                    </label>
                    <label>
                      <input
                        type="radio"
                        name="gPay"
                        value="khqr"
                        checked={gPayMethod === 'khqr'}
                        onChange={() => setGPayMethod('khqr')}
                        style={{ width: 'auto' }}
                      /> {t("KHQR")}
                    </label>

                    {gPayMethod === 'khqr' && (
                      <div id="gKhqr">
                        {qr.img ? (
                          <>
                            <img
                              id="gQrImg"
                              src={qr.img}
                              alt="KHQR code"
                              style={{ maxWidth: 240, width: '100%', margin: '0 auto', display: 'block', borderRadius: 12 }}
                            />
                            <div id="gQrName" style={{ textAlign: 'center', fontWeight: 600 }}>
                              {qr.name}
                            </div>
                            <p className="note" id="gQrNote">
                              {t(`Scan with your banking app and pay exactly ${both(gSubtotal * (1 + TAX))}, then upload the screenshot.`)}
                            </p>
                            <label className="note" htmlFor="gProof">
                              {t("Upload the payment screenshot (required)")}
                            </label>
                            <input
                              id="gProof"
                              type="file"
                              accept="image/*"
                              onChange={e => {
                                const f = e.target.files?.[0];
                                if (f) {
                                  shrinkImage(f, 720, 0.6, data => {
                                    setGProofData(data);
                                    setGProofMsg(t("Screenshot ready."));
                                  });
                                }
                              }}
                            />
                            <div className="note" id="gProofMsg" role="status">
                              {gProofMsg}
                            </div>
                          </>
                        ) : (
                          <p className="note" style={{ color: 'var(--hot)' }}>
                            {t("KHQR is not available yet. Please choose Cash on Delivery.")}
                          </p>
                        )}
                      </div>
                    )}
                  </div>

                  <input
                    id="gTel"
                    type="tel"
                    inputMode="numeric"
                    maxLength={14}
                    placeholder={t("Telephone (optional)")}
                    aria-label="Telephone (optional)"
                    style={{ fontSize: 16, marginTop: 10 }}
                    value={gTel}
                    onChange={e => setGTel(e.target.value)}
                  />
                  <input
                    id="gTgUser"
                    maxLength={32}
                    placeholder={t("Telegram username, e.g. @username (optional)")}
                    aria-label="Telegram username (optional)"
                    style={{ fontSize: 16, marginTop: 8 }}
                    value={gTgUser}
                    onChange={e => setGTgUser(e.target.value)}
                  />
                  <input
                    id="gNote"
                    maxLength={120}
                    placeholder={t("Note for the kitchen (optional)")}
                    aria-label="Note for the kitchen (optional)"
                    style={{ fontSize: 16, marginTop: 8 }}
                    value={gNote}
                    onChange={e => setGNote(e.target.value)}
                  />

                  <div className="note" id="gErr" role="alert" style={{ color: 'var(--hot)', minHeight: 20 }}>
                    {gErr && t(gErr)}
                  </div>

                  <div className="pay">
                    <button id="gBack" onClick={() => setGDlgOpen(false)}>{t("Back")}</button>
                    <button className="go" id="gPlace" onClick={handlePlaceGuestOrder}>{t("Place order")}</button>
                  </div>
                </div>
              ) : (
                <div id="gDone">
                  <h2 style={{ marginTop: 0 }}>
                    {t("Thank you! Your order was sent to the staff.")}
                  </h2>
                  <p className="note">{t("Order")} <b id="gOrderNo">{gDoneOrderNo}</b></p>
                  <button
                    className="go"
                    id="gNew"
                    style={{ width: '100%', padding: 14 }}
                    onClick={() => {
                      setGDlgOpen(false);
                      setGDoneOrderNo('');
                      const bytes = new Uint8Array(4);
                      crypto.getRandomValues(bytes);
                      const code = Array.from(bytes).map(b => CODE_CH[b % CODE_CH.length]).join('');
                      setGCode(code);
                      sessionStorage.setItem('pos-gcode', code);
                    }}
                  >
                    {t("New order")}
                  </button>
                </div>
              )}
            </dialog>
          )}

          {/* GUEST RECEIPT MODAL */}
          {gRcptOpen && activeMyGuestOrder?.receipt && (
            <dialog id="gRcpt" open style={{ width: 'min(460px, 94vw)', maxHeight: '90vh', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 45 }}>
              <h2 style={{ marginTop: 0 }}>{t("Your receipt")}</h2>
              <div id="gRcptBody" style={{ maxHeight: '60vh', overflow: 'auto' }}>
                <div className="rcpt">
                  <h3>{activeMyGuestOrder.receipt.name}</h3>
                  <p className="note">
                    {t(/\d/.test(activeMyGuestOrder.receipt.table) ? `Table ${activeMyGuestOrder.receipt.table}` : activeMyGuestOrder.receipt.table)} - {activeMyGuestOrder.receipt.time} - {t("paid by")} {t(activeMyGuestOrder.receipt.method)}
                  </p>
                  <div className="r3">
                    <b>{t("Item")}</b>
                    <b>{t("Qty")}</b>
                    <b>{t("Price")}</b>
                  </div>
                  {activeMyGuestOrder.receipt.items?.map((it: any) => (
                    <div className="r3" key={it[0]}>
                      <span>{t(it[0])}</span>
                      <span>{it[1]}</span>
                      <span>{m(it[2])}</span>
                    </div>
                  ))}
                  <div className="row">
                    <span>{t("Subtotal")}</span>
                    <span>{m(activeMyGuestOrder.receipt.sub)}</span>
                  </div>
                  <div className="row">
                    <span>{t("Tax (8%)")}</span>
                    <span>{m(activeMyGuestOrder.receipt.tax)}</span>
                  </div>
                  <div className="row">
                    <span>{t("Tip")}</span>
                    <span>{m(activeMyGuestOrder.receipt.tip)}</span>
                  </div>
                  <div className="row grand">
                    <span>{t("Grand total")}</span>
                    <span>{m(activeMyGuestOrder.receipt.tot)}</span>
                  </div>
                  <div className="row">
                    <span>{t("Grand total in riel")}</span>
                    <span>{khr(activeMyGuestOrder.receipt.tot, activeMyGuestOrder.receipt.rate)}</span>
                  </div>
                </div>
              </div>
              <button
                className="go"
                id="gRcptOk"
                style={{ width: '100%', padding: 14, marginTop: 10 }}
                onClick={() => setGRcptOpen(false)}
              >
                {t("OK")}
              </button>
            </dialog>
          )}

          {/* GUEST RE-UPLOAD SCREENSHOT DIALOG */}
          {gReDlgOpen && (
            <dialog id="gRe" open style={{ width: 'min(420px, 94vw)', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 45 }}>
              <h2 style={{ marginTop: 0 }}>{t("Upload new screenshot")}</h2>
              <input
                id="gReFile"
                type="file"
                accept="image/*"
                onChange={e => {
                  const f = e.target.files?.[0];
                  if (f) {
                    shrinkImage(f, 720, 0.6, data => setGReProof(data));
                  }
                }}
              />
              <div className="note" id="gReMsg" role="status">
                {gReProof ? t("Screenshot ready.") : ""}
              </div>
              <div className="pay">
                <button id="gReCancel" onClick={() => setGReDlgOpen(false)}>{t("Cancel")}</button>
                <button
                  className="go"
                  id="gReSend"
                  onClick={async () => {
                    if (!gReProof) {
                      showToast("Please upload your payment screenshot.");
                      return;
                    }
                    const updated = guestOrders.map(o => o.id === gMyActiveId ? { ...o, proof: gReProof, payStatus: 'pending', payNote: '' } : o);
                    setGuestOrders(updated);
                    localStorage.setItem('pos-gorders', JSON.stringify(updated));
                    try {
                      await updateDoc(doc(db, 'guestorders', gMyActiveId), { proof: gReProof, payStatus: 'pending', payNote: '' });
                    } catch (e) {}
                    setGReDlgOpen(false);
                  }}
                >
                  {t("Send")}
                </button>
              </div>
            </dialog>
          )}
        </div>
      )}

      {/* GUEST ORDERS INBOX MODAL */}
      {goDlgOpen && (
        <dialog id="goDlg" open style={{ width: 'min(520px, 94vw)', maxHeight: '90vh', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40, overflow: 'auto' }}>
          <h2 style={{ marginTop: 0 }}>{t("Guest orders")}</h2>
          <div id="goList" style={{ maxHeight: '62vh', overflow: 'auto' }}>
            {newGuestOrders.length > 0 ? (
              newGuestOrders.map(o => {
                const tt = o.items.reduce((acc, i) => acc + getPrice(i[0]) * (i[1] || 0), 0) * (1 + TAX);
                const assigned = gAssignTable[o.id] || (o.table && TABLES.includes(o.table) ? o.table : freeTable);
                const reason = gReasons[o.id] || '';

                return (
                  <div className="rec" key={o.id}>
                    <b>Order #{o.code || o.id.slice(-4).toUpperCase()}</b> - {o.time}
                    <ul style={{ margin: '4px 0', paddingLeft: 20 }}>
                      {o.items.map(i => (
                        <li key={i[0]}>{t(i[0])} x{i[1]}</li>
                      ))}
                    </ul>
                    <div className="row">
                      <span>{t("Total")}</span>
                      <b>{m(tt)} / {kh(tt)}</b>
                    </div>
                    {o.tel && <div className="note">{t("Phone")} {o.tel}</div>}
                    {o.tgUser && <div className="note">Telegram: {o.tgUser}</div>}
                    {o.note && <div className="note">{o.note}</div>}

                    {/* Payment verification */}
                    <div className="note" style={{ marginTop: 4 }}>
                      <b>{t("Payment")}</b> <span>{o.pay === 'khqr' ? 'KHQR' : t("Cash on Delivery")}</span>
                    </div>
                    {o.pay === 'khqr' && o.proof && (
                      <img
                        src={o.proof}
                        alt="KHQR proof"
                        style={{ maxHeight: 180, maxWidth: '100%', borderRadius: 10, cursor: 'zoom-in', display: 'block', margin: '4px 0' }}
                        onClick={() => setProofImg(o.proof || '')}
                      />
                    )}
                    {o.pay === 'khqr' && (
                      <>
                        <div className="note">
                          <b>
                            {o.payStatus === 'matched' ? t("Payment matched") : o.payStatus === 'mismatch' ? `${t("Not matching")}${o.payNote ? ': ' + o.payNote : ''}` : t("Waiting for your check")}
                          </b>
                        </div>
                        {o.payStatus !== 'matched' && (
                          <div className="pay" style={{ margin: '4px 0' }}>
                            <button
                              onClick={async () => {
                                if (!reason) {
                                  showToast(t("Write a reason first."));
                                  return;
                                }
                                const updated = guestOrders.map(item => item.id === o.id ? { ...item, payStatus: 'mismatch', payNote: reason } : item);
                                setGuestOrders(updated);
                                localStorage.setItem('pos-gorders', JSON.stringify(updated));
                                try {
                                  await updateDoc(doc(db, 'guestorders', o.id), { payStatus: 'mismatch', payNote: reason });
                                } catch (e) {}
                                showToast(t("Marked as not matching"));
                              }}
                            >
                              {t("Not matching")}
                            </button>
                            <button
                              className="go"
                              onClick={async () => {
                                const updated = guestOrders.map(item => item.id === o.id ? { ...item, payStatus: 'matched', payNote: reason } : item);
                                setGuestOrders(updated);
                                localStorage.setItem('pos-gorders', JSON.stringify(updated));
                                try {
                                  await updateDoc(doc(db, 'guestorders', o.id), { payStatus: 'matched', payNote: reason });
                                } catch (e) {}
                                showToast(t("Payment marked as matching"));
                              }}
                            >
                              {t("Payment matches")}
                            </button>
                          </div>
                        )}
                      </>
                    )}

                    <div className="note" style={{ marginTop: 6 }}>{t("Assign to table")}</div>
                    <select
                      value={assigned}
                      onChange={e => setGAssignTable(prev => ({ ...prev, [o.id]: e.target.value }))}
                      style={{ width: '100%', fontSize: 16, padding: 8, border: '1px solid var(--line)', borderRadius: 10, background: 'var(--bg)', color: 'var(--ink)' }}
                    >
                      {TABLES.map(tName => (
                        <option key={tName} value={tName}>
                          {/\d/.test(tName) ? `${t("Table")} ${tName}` : t(tName)}
                        </option>
                      ))}
                    </select>

                    <input
                      value={reason}
                      onChange={e => setGReasons(prev => ({ ...prev, [o.id]: e.target.value }))}
                      maxLength={120}
                      placeholder={t("Note / reason (required to decline)")}
                      style={{ width: '100%', fontSize: 16, padding: 8, margin: '6px 0 0', border: '1px solid var(--line)', borderRadius: 10, background: 'var(--bg)', color: 'var(--ink)' }}
                    />

                    <div className="pay" style={{ marginTop: 6 }}>
                      <button
                        style={{ color: 'var(--hot)', borderColor: 'var(--hot)' }}
                        onClick={async () => {
                          if (!reason) {
                            showToast(t("Write a reason to decline."));
                            return;
                          }
                          const updated = guestOrders.map(item => item.id === o.id ? { ...item, status: 'declined' as const, reason, decidedAt: Date.now() } : item);
                          setGuestOrders(updated);
                          localStorage.setItem('pos-gorders', JSON.stringify(updated));
                          try {
                            await updateDoc(doc(db, 'guestorders', o.id), { status: 'declined', reason, decidedAt: Date.now() });
                          } catch (e) {}
                          showToast("Order declined");
                        }}
                      >
                        {t("Decline")}
                      </button>
                      <button
                        className="go"
                        onClick={async () => {
                          if (o.pay === 'khqr' && o.payStatus !== 'matched') {
                            showToast(t("Check the KHQR payment first."));
                            return;
                          }

                          // Add items to assigned table's cart
                          setOrders(prev => {
                            const cur = { ...(prev[assigned] || {}) };
                            o.items.forEach(([n, q]) => {
                              cur[n] = (cur[n] || 0) + (Number(q) || 0);
                            });
                            return { ...prev, [assigned]: cur };
                          });

                          if (o.tel && /^\d{9,10}$/.test(o.tel)) {
                            setGTels(prev => ({ ...prev, [assigned]: o.tel! }));
                          }
                          if (o.pay === 'khqr') {
                            setGPayPref(prev => ({ ...prev, [assigned]: 'KHQR' }));
                          }

                          setGOrdIds(prev => ({ ...prev, [assigned]: [...(prev[assigned] || []), o.id] }));

                          // Update order status
                          const updated = guestOrders.map(item => item.id === o.id ? { ...item, status: 'confirmed' as const, table: assigned, reason, decidedAt: Date.now() } : item);
                          setGuestOrders(updated);
                          try {
                            localStorage.setItem('pos-gorders', JSON.stringify(updated));
                            await updateDoc(doc(db, 'guestorders', o.id), { status: 'confirmed', table: assigned, reason, decidedAt: Date.now() });
                          } catch (e) {}

                          setTable(assigned);
                          setGoDlgOpen(false);
                          showToast(`Confirmed - added to the cart for ${/\d/.test(assigned) ? `table ${assigned}` : assigned}`);
                        }}
                      >
                        {t("Confirm")}
                      </button>
                    </div>
                  </div>
                );
              })
            ) : (
              <p className="empty">{t("No new guest orders.")}</p>
            )}

            {recentGuestOrders.length > 0 && (
              <>
                <h3 style={{ margin: '14px 0 4px' }}>{t("Recent decisions")}</h3>
                {recentGuestOrders.map(o => (
                  <div className="rec" key={o.id}>
                    <b>Order #{o.code || o.id.slice(-4).toUpperCase()}</b> - {o.time} - <b>{t(o.status === 'confirmed' ? "Confirmed" : "Declined")}</b>
                    {o.reason && <div className="note">{o.reason}</div>}
                  </div>
                ))}
              </>
            )}
          </div>
          <button id="goClose" style={{ width: '100%', marginTop: 10 }} onClick={() => setGoDlgOpen(false)}>
            {t("Close")}
          </button>
        </dialog>
      )}

      {/* PROOF ZOOM MODAL */}
      {proofImg && (
        <dialog id="proofDlg" open style={{ width: 'min(520px, 94vw)', maxHeight: '92vh', overflow: 'auto', display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 50 }}>
          <img id="proofImg" src={proofImg} alt="KHQR payment screenshot" style={{ width: '100%', borderRadius: 10 }} />
          <button id="proofClose" style={{ width: '100%', marginTop: 8 }} onClick={() => setProofImg('')}>
            {t("Close")}
          </button>
        </dialog>
      )}

      {/* PIN VERIFICATION MODAL */}
      {pinDlgOpen && (
        <dialog id="pinDlg" open style={{ display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 45 }}>
          <h2 style={{ marginTop: 0 }}>{t("Enter PIN")}</h2>
          <input
            id="pinIn"
            type="password"
            inputMode="numeric"
            maxLength={8}
            autoComplete="off"
            aria-label="Enter PIN"
            value={pinInput}
            onChange={e => {
              setPinInput(e.target.value);
              setPinErr('');
            }}
          />
          <div className="note" id="pinErr" role="alert" style={{ color: 'var(--hot)', minHeight: 20 }}>
            {pinErr && t(pinErr)}
          </div>
          <div className="pay">
            <button id="pinCancel" onClick={() => setPinDlgOpen(false)}>{t("Cancel")}</button>
            <button
              className="go"
              id="pinOk"
              onClick={async () => {
                const h = await hashPin(pinInput.trim());
                const ok = pinTarget === 'admin' ? h === pins.A : (h === pins.S || h === pins.A);
                if (ok) {
                  setPinDlgOpen(false);
                  setRole(pinTarget);
                } else {
                  setPinErr('Incorrect PIN.');
                }
              }}
            >
              {t("OK")}
            </button>
          </div>
        </dialog>
      )}

      {/* ACCESS & PIN SETTINGS MODAL */}
      {accDlgOpen && (
        <dialog id="accDlg" open style={{ display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40 }}>
          <h2 style={{ marginTop: 0 }}>{t("Access")}</h2>
          <div className="mform">
            <label className="note" htmlFor="pinAdmin">{t("Admin PIN")}</label>
            <div style={{ display: 'flex', gap: 6 }}>
              <input
                id="pinAdmin"
                type="password"
                inputMode="numeric"
                maxLength={8}
                placeholder="4-8 digits"
                style={{ fontSize: 16 }}
                value={adminPinIn}
                onChange={e => setAdminPinIn(e.target.value)}
              />
              <button
                id="accClrA"
                onClick={async () => {
                  setPins(prev => {
                    const next = { ...prev, A: '' };
                    localStorage.setItem('pos-pinA', '');
                    return next;
                  });
                  try {
                    await setDoc(doc(db, 'pos', 'settings'), { pinA: '' }, { merge: true });
                  } catch (e) {}
                  setAccMsg("Admin PIN removed.");
                }}
              >
                {t("Remove PIN")}
              </button>
            </div>

            <label className="note" htmlFor="pinStaff">{t("Staff PIN")}</label>
            <div style={{ display: 'flex', gap: 6 }}>
              <input
                id="pinStaff"
                type="password"
                inputMode="numeric"
                maxLength={8}
                placeholder="4-8 digits"
                style={{ fontSize: 16 }}
                value={staffPinIn}
                onChange={e => setStaffPinIn(e.target.value)}
              />
              <button
                id="accClrS"
                onClick={async () => {
                  setPins(prev => {
                    const next = { ...prev, S: '' };
                    localStorage.setItem('pos-pinS', '');
                    return next;
                  });
                  try {
                    await setDoc(doc(db, 'pos', 'settings'), { pinS: '' }, { merge: true });
                  } catch (e) {}
                  setAccMsg("Staff PIN removed.");
                }}
              >
                {t("Remove PIN")}
              </button>
            </div>

            <div className="note" id="accMsg" role="status">
              {accMsg}
            </div>
            <p className="note" style={{ margin: 0 }}>
              PINs only keep staff and guests out of the wrong page on this screen.
            </p>
          </div>
          <div className="pay">
            <button id="accClose" onClick={() => setAccDlgOpen(false)}>{t("Close")}</button>
            <button
              className="go"
              id="accSave"
              onClick={async () => {
                const a = adminPinIn.trim();
                const b = staffPinIn.trim();
                const valid = (v: string) => !v || /^\d{4,8}$/.test(v);
                if (!valid(a) || !valid(b)) {
                  setAccMsg("A PIN must be 4 to 8 digits.");
                  return;
                }
                const newPins = { ...pins };
                if (a) newPins.A = await hashPin(a);
                if (b) newPins.S = await hashPin(b);
                setPins(newPins);
                localStorage.setItem('pos-pinA', newPins.A);
                localStorage.setItem('pos-pinS', newPins.S);
                try {
                  await setDoc(doc(db, 'pos', 'settings'), { pinA: newPins.A, pinS: newPins.S }, { merge: true });
                } catch (e) {}
                setAccMsg("Saved.");
                setAdminPinIn('');
                setStaffPinIn('');
              }}
            >
              {t("Save")}
            </button>
          </div>
        </dialog>
      )}

      {/* DATABASE & BACKUP MODAL */}
      {dbDlgOpen && (
        <dialog id="dbDlg" open style={{ display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40 }}>
          <h2 style={{ marginTop: 0 }}>{t("Database")}</h2>
          <p id="dbStatus" style={{ margin: '0 0 6px' }}>
            <b style={{ color: cloudConnected ? '#1f7a4d' : 'inherit' }}>
              {cloudConnected ? "🟢 Connected to Firebase Firestore." : "Connecting to Firebase Firestore..."}
            </b>
            <br />
            Menu, settings, ad picture, and sales are synced in real-time across all devices.
          </p>
          <p className="note" id="dbCount" style={{ margin: '0 0 10px' }}>
            {menu.reduce((acc: number, g: any) => acc + g.i.length, 0)} menu items - {sales.length} sales loaded
          </p>
          <div className="mform">
            <button
              id="dbUpload"
              onClick={async () => {
                setDbUploadMsg("Uploading data to Firebase...");
                try {
                  // Upload menu items
                  for (const grp of menu) {
                    for (const item of grp.i) {
                      const id = 'm_' + encodeURIComponent(item[0]).replace(/%/g, '_');
                      await setDoc(doc(db, 'items', id), {
                        cat: grp.c,
                        name: item[0],
                        price: item[1],
                        img: item[2],
                        promo: !!item[3],
                        desc: item[4] || '',
                        o: Date.now()
                      });
                    }
                  }
                  // Upload sales
                  for (const s of sales) {
                    if (s.id) {
                      await setDoc(doc(db, 'sales', s.id), s);
                    }
                  }
                  // Upload settings
                  await setDoc(doc(db, 'pos', 'settings'), {
                    name: rname,
                    rate,
                    promoMsg,
                    promoLink,
                    qrName: qr.name
                  }, { merge: true });

                  setDbUploadMsg("Upload to Firebase complete!");
                  showToast("Upload to Firebase complete!");
                } catch (e: any) {
                  setDbUploadMsg("Upload failed: " + (e?.message || 'Error'));
                }
              }}
            >
              {t("Upload this device's saved data")}
            </button>
            {dbUploadMsg && <div className="note" style={{ color: 'var(--acc)', fontWeight: 600 }}>{dbUploadMsg}</div>}

            <button
              id="dbBackup"
              onClick={() => {
                const backupData = {
                  exported: new Date().toISOString(),
                  restaurant: rname,
                  menu,
                  sales,
                  settings: { promoMsg, promoLink, rate, qr }
                };
                const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `pos-backup-${new Date().toLocaleDateString('en-CA')}.json`;
                document.body.appendChild(a);
                a.click();
                a.remove();
                setTimeout(() => URL.revokeObjectURL(url), 4000);
              }}
            >
              {t("Download backup (JSON)")}
            </button>
            <p className="note" style={{ margin: 0 }}>
              Sales records include guest telephone numbers. The Telegram bot token is never stored in public backups.
            </p>
          </div>
          <button id="dbClose" style={{ width: '100%', marginTop: 8 }} onClick={() => setDbDlgOpen(false)}>
            {t("Close")}
          </button>
        </dialog>
      )}

      {/* TELEGRAM BOT MODAL */}
      {tgDlgOpen && (
        <dialog id="tgDlg" open style={{ display: 'block', position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', zIndex: 40 }}>
          <h2 style={{ marginTop: 0 }}>{t("Telegram bot")}</h2>
          <div className="mform">
            <label className="note" htmlFor="tgToken">{t("Bot token (from @BotFather)")}</label>
            <input
              id="tgToken"
              type="password"
              autoComplete="off"
              placeholder="123456:ABC..."
              style={{ fontSize: 16 }}
              value={tg.token}
              onChange={e => setTg(prev => ({ ...prev, token: e.target.value.trim() }))}
            />
            <label className="note" htmlFor="tgChat">{t("Chat ID (a number, or @channelname)")}</label>
            <input
              id="tgChat"
              placeholder="-1001234567890"
              style={{ fontSize: 16 }}
              value={tg.chat}
              onChange={e => setTg(prev => ({ ...prev, chat: e.target.value.trim() }))}
            />

            <label className="note">
              <input
                type="checkbox"
                id="tgOrderAlert"
                style={{ width: 'auto' }}
                checked={tg.orderAlert}
                onChange={e => setTg(prev => ({ ...prev, orderAlert: e.target.checked }))}
              /> {t("Alert to Telegram admin group when guest orders")}
            </label>

            <label className="note">
              <input
                type="checkbox"
                id="tgAuto"
                style={{ width: 'auto' }}
                checked={tg.auto}
                onChange={e => setTg(prev => ({ ...prev, auto: e.target.checked }))}
              /> {t("Send the receipt automatically after each payment")}
            </label>
            <p className="note" style={{ margin: 0 }}>
              The token is saved only in this browser.
            </p>
          </div>
          <div className="pay" style={{ gridTemplateColumns: '1fr 1fr' }}>
            <button
              id="tgAlertTest"
              onClick={async () => {
                localStorage.setItem('pos-tg', JSON.stringify(tg));
                showToast("Sending test order alert...");
                const testOrder: GuestOrder = {
                  id: 'g' + Date.now().toString(36),
                  code: '758e0885',
                  table: '',
                  pay: 'cod',
                  payStatus: 'cod',
                  items: [
                    ['Iced Latte', 3, 3.50]
                  ],
                  total: 10.50,
                  rate,
                  tel: '070202020',
                  tgUser: '@RaksmeyIXIII',
                  note: 'Less Sweet',
                  ts: Date.now(),
                  date: new Date().toLocaleDateString('en-CA'),
                  time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                  status: 'new'
                };
                const res = await sendTelegramOrderAlert(testOrder);
                showToast(res?.ok ? "Test order alert sent to Telegram!" : res?.error);
              }}
            >
              {t("Send test order alert")}
            </button>
            <button
              id="tgTest"
              onClick={async () => {
                localStorage.setItem('pos-tg', JSON.stringify(tg));
                showToast("Sending...");
                const res = await tgSend(`${rname}: test message from the POS.`);
                showToast(res.ok ? "Test message sent. Check your Telegram chat." : res.error);
              }}
            >
              {t("Send test message")}
            </button>
          </div>
          <div className="pay" style={{ gridTemplateColumns: '1fr', marginTop: 8 }}>
            <button
              className="go"
              id="tgSave"
              onClick={() => {
                localStorage.setItem('pos-tg', JSON.stringify(tg));
                showToast("Saved Telegram settings.");
              }}
            >
              {t("Save")}
            </button>
          </div>
          <button id="tgClose" style={{ width: '100%', marginTop: 8 }} onClick={() => setTgDlgOpen(false)}>
            {t("Close")}
          </button>
        </dialog>
      )}

      {/* TOAST OVERLAY */}
      {toastMsg && (
        <div id="toast" role="status" style={{ position: 'fixed', left: '50%', bottom: 20, transform: 'translateX(-50%)', background: 'var(--ink)', color: 'var(--bg)', padding: '10px 16px', borderRadius: 10, zIndex: 50, maxWidth: '92vw' }}>
          {t(toastMsg)}
        </div>
      )}
    </>
  );
}
