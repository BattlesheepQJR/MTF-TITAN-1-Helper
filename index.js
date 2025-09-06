const { Client, GatewayIntentBits, SlashCommandBuilder } = require('discord.js');
const { google } = require('googleapis');
const { MessageFlags } = require('discord.js');
const { InteractionResponseFlags } = require('discord.js');
const fs = require('fs');
const path = require('path');
const express = require('express')
const app = express()
const port = process.env.PORT || 4000

var currentTask = 1;

// ======================= DISCORD BOT =======================

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMembers, 
  ],
});

client.once('clientReady', () => {
  console.log(`┌[ALL] [BOT LOGIN] ${client.user.displayName} Has Logged In.`);
});

client.on('clientReady', async () => {
  const data = new SlashCommandBuilder()
    .setName('rankchange')
    .setDescription('Change a soldier rank')
    .addStringOption(option =>
      option.setName('action')
        .setDescription('Promote, Demote, Transfer')
        .setRequired(true)
        .addChoices(
          { name: 'Promote', value: 'Promotion' },
          { name: 'Demote', value: 'Demotion' },
          { name: 'Transfer', value: 'Transfer' }
        ))
    .addUserOption(option =>
      option.setName('user')
        .setDescription('Discord user to change rank of')
        .setRequired(true))
    .addStringOption(option =>
      option.setName('note')
        .setDescription('A Description / Reason Of The Rank Change.')
        .setRequired(false))
    .addStringOption(option =>
      option.setName('rank')
        .setDescription('Rank To Be Given To The User.')
        .setRequired(false)
        .addChoices(
          { name: 'General', value: 'General' },
          { name: 'Lieutenant General', value: 'Lieutenant General' },
          { name: 'Major General', value: 'Major General' },
          { name: 'Brigadier General', value: 'Brigadier General' },
          { name: 'Colonel', value: 'Colonel' },
          { name: 'Lieutenant Colonel', value: 'Lieutenant Colonel' },
          { name: 'Major', value: 'Major' },
          { name: 'Captain', value: 'Captain' },
          { name: 'First Lieutenant', value: 'First Lieutenant' },
          { name: 'Second Lieutenant', value: 'Second Lieutenant' },
          { name: 'Regimental Sergeant Major', value: 'Regimental Sergeant Major' },
          { name: 'Sergeant Major', value: 'Sergeant Major' },
          { name: 'First Sergeant', value: 'First Sergeant' },
          { name: 'Master Sergeant', value: 'Master Sergeant' },
          { name: 'Sergeant First Class', value: 'Sergeant First Class' },
          { name: 'Staff Sergeant', value: 'Staff Sergeant' },
          { name: 'Sergeant', value: 'Sergeant' },
          { name: 'Corporal', value: 'Corporal' },
          { name: 'Lance Corporal', value: 'Lance Corporal' },
          { name: 'Private First Class', value: 'Private First Class' },
          { name: 'Private Second Class', value: 'Private Second Class' },
          { name: 'Private', value: 'Private' },
        ));
  await client.application.commands.create(data);
});

// Handle command
client.on('interactionCreate', async interaction => {
  if (!interaction.isChatInputCommand()) return;
  if (interaction.commandName !== 'rankchange') return;

  const action = interaction.options.getString('action');
  const user = interaction.options.getUser('user');
  var rank = interaction.options.getString('rank');
  const note = interaction.options.getString('note');
  const co = interaction.user;

  const username = user.username;
  const displayName = user.displayName;
  const coUsername = co.username;
  const coDisplayName = co.displayName;
  const owner = await interaction.guild.fetchOwner()
  const ownerUsername = owner.user.username

  var action2

  if (action === "Promotion") {
    action2 = "Promote"
  }
  if (action === "Demotion") {
    action2 = "Demote"
  }
  if (action === "Transfer") {
    action2 = "Transfer"
  }
  
  if (username === coUsername && coUsername !== ownerUsername) {
    return interaction.reply({
      content: `You Cannot ${action2} Yourself!`
    });
  }

  if (username === ownerUsername && coUsername !== ownerUsername) {
    return interaction.reply({
      content: `You Cannot ${action2} The Commanding General!`
    });
  }

  if (rank === null && action === "Promotion") {
    const resRank = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Member Data!B:O'
    });

    const rowsRank = resRank.data.values;
    if (!rowsRank || rowsRank.length === 0) {
      return "❌ No data found in Member Data.";
    }

    const headersRank = rowsRank[1];
    const discordIndexRank = headersRank.indexOf('Discord');
    
    const targetRowIndexrowsRank = rowsRank.findIndex(r => r[discordIndexRank] == username);
    const userPrevRank = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `Member Data!B${targetRowIndexrowsRank+1}:H${targetRowIndexrowsRank+1}`
    });
    const userPrevRankValues = userPrevRank.data.values;
    const userPrevRankValue = userPrevRankValues[0][4];
    
    if (userPrevRankValues[0][6] === 'TRUE') {
      return interaction.reply({
        content: `<@${user.id}> Is Retired!.`,
        allowedMentions: { parse: [] }
      });
    }

    const resRank1 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Lookups!B:O'
    });

    const rowsRank1 = resRank1.data.values;
    if (!rowsRank1 || rowsRank1.length === 0) {
      return "❌ No data found in Member Data.";
    }

    const headersRank1 = rowsRank1[1];
    const rankIndexRank = headersRank1.indexOf('Rank');

    const targetRowIndexrowsRank1 = rowsRank1.findIndex(r => r[rankIndexRank] == userPrevRankValue);
    const userNewRank1 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `Lookups!B${targetRowIndexrowsRank1+2}:F${targetRowIndexrowsRank1+2}`
    });
    const userNewRankValues1 = userNewRank1.data.values;
    const userNewRankValue1 = userNewRankValues1[0][2];
    rank = userNewRankValue1
  } else if (rank === null && action === "Demotion") {
    const resRank = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Member Data!B:O'
    });

    const rowsRank = resRank.data.values;
    if (!rowsRank || rowsRank.length === 0) {
      return "❌ No data found in Member Data.";
    }

    const headersRank = rowsRank[1];
    const discordIndexRank = headersRank.indexOf('Discord');

    const targetRowIndexrowsRank = rowsRank.findIndex(r => r[discordIndexRank] == username);
    const userPrevRank = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `Member Data!B${targetRowIndexrowsRank+1}:H${targetRowIndexrowsRank+1}`
    });
    const userPrevRankValues = userPrevRank.data.values;
    const userPrevRankValue = userPrevRankValues[0][4];

    if (userPrevRankValues[0][6] === 'TRUE') {
      return interaction.reply({
        content: `<@${user.id}> Is Retired!.`,
        allowedMentions: { parse: [] }
      });
    }

    const resRank1 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Lookups!B:O'
    });

    const rowsRank1 = resRank1.data.values;
    if (!rowsRank1 || rowsRank1.length === 0) {
      return "❌ No data found in Member Data.";
    }

    const headersRank1 = rowsRank1[1];
    const rankIndexRank = headersRank1.indexOf('Rank');

    const targetRowIndexrowsRank1 = rowsRank1.findIndex(r => r[rankIndexRank] == userPrevRankValue);
    const userNewRank1 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `Lookups!B${targetRowIndexrowsRank1}:F${targetRowIndexrowsRank1}`
    });
    const userNewRankValues1 = userNewRank1.data.values;
    const userNewRankValue1 = userNewRankValues1[0][2];
    rank = userNewRankValue1
  } else if (rank === null && action === "Transfer") {
    return interaction.reply({
      content: `You Need To Include A Rank To Transfer <@${user.id}>.`,
      allowedMentions: { parse: [] }
    });
  }

  console.log(`├[${currentTask}]`)
  console.log(`│ ├[RANK CHANGE - Request] ${action2} ${displayName} To ${rank}.`)

  if (rank === 'Commanding General' || rank === 'Foundation Command') {
    console.log(`│ └[RANK ALLOTMENT] ${rank} Is Disabled`)
    return interaction.reply({
      content: `Cannot Promote <@${user.id}> To ${rank} As It Is Disabled.`,
      allowedMentions: { parse: [] }
    });
  }

  if (action === "Promotion" && rank) {
    const res0 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Lookups!B:O'
    });

    const rows0 = res0.data.values;
    if (!rows0 || rows0.length === 0) {
      return "❌ No data found in Rank Change.";
    }

    const headers0 = rows0[1];
    const rankIndex0 = headers0.indexOf('Rank');

    const targetRowIndex0 = rows0.findIndex(r => r[rankIndex0] == rank);

    const rank0 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `Lookups!F${targetRowIndex0+1}:H${targetRowIndex0+1}`
    });
    const rank0Values = rank0.data.values;

    if (rank0Values[0][2] === 'FALSE') {
      console.log(`│ └[RANK ALLOTMENT] Member Count Of "${rank}" Is At The Max`)
      return interaction.reply({
        content: `Cannot Promote <@${user.id}> To ${rank} As The Max Of **${rank0Values[0][0]}** Has Already Been Achieved.`,
        allowedMentions: { parse: [] }
      });
    }
  } else if (action === "Demotion" && rank) {
    const res0 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Lookups!B:O'
    });

    const rows0 = res0.data.values;
    if (!rows0 || rows0.length === 0) {
      return "❌ No data found in Rank Change.";
    }

    const headers0 = rows0[1];
    const rankIndex0 = headers0.indexOf('Rank');

    const targetRowIndex0 = rows0.findIndex(r => r[rankIndex0] == rank);

    const rank0 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `Lookups!F${targetRowIndex0+1}:H${targetRowIndex0+1}`
    });
    const rank0Values = rank0.data.values;

    if (rank0Values[0][2] === 'FALSE') {
      console.log(`│ └[RANK ALLOTMENT] Member Count Of "${rank}" Is At The Max`)
      return interaction.reply({
        content: `Cannot Demote <@${user.id}> To ${rank} As The Max Of ${rank0Values[0][0]} Has Already Been Achieved.`,
        allowedMentions: { parse: [] }
      });
    }
  } else if (action === "Transfer" && rank) {
    const res0 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Lookups!B:O'
    });

    const rows0 = res0.data.values;
    if (!rows0 || rows0.length === 0) {
      return "❌ No data found in Rank Change.";
    }

    const headers0 = rows0[1];
    const rankIndex0 = headers0.indexOf('Rank');

    const targetRowIndex0 = rows0.findIndex(r => r[rankIndex0] == rank);

    const rank0 = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `Lookups!F${targetRowIndex0+1}:H${targetRowIndex0+1}`
    });
    const rank0Values = rank0.data.values;

    if (rank0Values[0][2] === 'FALSE') {
      console.log(`│ └[RANK ALLOTMENT] Member Count Of "${rank}" Is At The Max`)
      return interaction.reply({
        content: `Cannot Transfer <@${user.id}> To ${rank} As The Max Of ${rank0Values[0][0]} Has Already Been Achieved.`,
        allowedMentions: { parse: [] }
      });
    }
  }

  await interaction.deferReply({});
  await interaction.editReply({
    content: "⏳ Updating sheet..." 
  });
  try {
    const result = await updateRank(action, username, rank, coUsername, note, user);
    await interaction.editReply({
      content: result,
      allowedMentions: { parse: [] }
    });
  } catch (err) {
    console.error(err);
    await interaction.editReply({
      content: "❌ An error occurred while updating the sheet.",
    })
  }
  console.log(`│ └[RANK CHANGE - ACTION] ${displayName} Is Now A ${rank}.`)
  currentTask = currentTask + 1
});

client.login(process.env['TOKEN']);

// ======================= GOOGLE SHEETS =======================

const credentials = JSON.parse(process.env['CREDENTIALS']);

const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});
const sheets = google.sheets({ version: 'v4', auth });
const SHEET_ID = process.env['SHEET_ID']
const RANK_CHANGE_SHEET_ID = process.env['RANK_CHANGE_SHEET_ID']

/**
 * Update rank in Member Data sheet
 */
async function updateRank(action, username, rank, coUsername, note, user) {
  // Fetch all rows
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'Rank Change!B3:Q3'
  });

  const res1 = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'Member Data!B:O'
  });

  const rows = res.data.values;
  if (!rows || rows.length === 0) {
    return "❌ No data found in Rank Change.";
  }

  const rows1 = res1.data.values;
  if (!rows1 || rows1.length === 0) {
    return "❌ No data found in Member Data.";
  }

  // Header positions
  const headers1 = rows1[1];
  const discordIndex = headers1.indexOf('Discord');

  const targetRowIndex1 = rows1.findIndex(r => r[discordIndex] == username);
  const targetRowIndex2 = rows1.findIndex(r => r[discordIndex] == coUsername)

  const userId = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `Member Data!B${targetRowIndex1+1}`
  });
  const userIdValues = userId.data.values;
  const userIdValue = userIdValues[0][0];

  const coUserId = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `Member Data!B${targetRowIndex2+1}`
  });
  const coUserIdValues = coUserId.data.values;
  const coUserIdValue = coUserIdValues[0][0];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: RANK_CHANGE_SHEET_ID,          // The numeric ID of the sheet (0 = usually the first sheet)
              dimension: "ROWS",   // "ROWS" or "COLUMNS"
              startIndex: 5,       // Zero-based index → row 5 is index 4
              endIndex: 6
            },
            inheritFromBefore: true // Whether new row copies formatting from the row before
          }
        }
      ]
    }
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: {
              sheetId: RANK_CHANGE_SHEET_ID, // not the name, but the numeric sheetId
              startRowIndex: 5,  // Row 6 (zero-based)
              endRowIndex: 6,
              startColumnIndex: 1, // B (zero-based)
              endColumnIndex: 5,   // F
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 0.235,
                  green: 0.251,
                  blue: 0.263,
                },
              },
            },
            fields: "userEnteredFormat.backgroundColor",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId: RANK_CHANGE_SHEET_ID, // not the name, but the numeric sheetId
              startRowIndex: 5,  // Row 6 (zero-based)
              endRowIndex: 6,
              startColumnIndex: 6, // B (zero-based)
              endColumnIndex: 11,   // F
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 0.235,
                  green: 0.251,
                  blue: 0.263,
                },
              },
            },
            fields: "userEnteredFormat.backgroundColor",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId: RANK_CHANGE_SHEET_ID, // not the name, but the numeric sheetId
              startRowIndex: 5,  // Row 6 (zero-based)
              endRowIndex: 6,
              startColumnIndex: 12, // B (zero-based)
              endColumnIndex: 13,   // F
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 0.235,
                  green: 0.251,
                  blue: 0.263,
                },
              },
            },
            fields: "userEnteredFormat.backgroundColor",
          },
        },
        {
          repeatCell: {
            range: {
              sheetId: RANK_CHANGE_SHEET_ID, // not the name, but the numeric sheetId
              startRowIndex: 5,  // Row 6 (zero-based)
              endRowIndex: 6,
              startColumnIndex: 14, // B (zero-based)
              endColumnIndex: 16,   // F
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 0.235,
                  green: 0.251,
                  blue: 0.263,
                },
              },
            },
            fields: "userEnteredFormat.backgroundColor",
          },
        },
      ]
    }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!H5`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [['=ARRAYFORMULA({"INS"; IF($G$6:G<>"", VLOOKUP($G$6:G,Lookups!$D$4:$E$27, 2, False), "")})']] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!J5`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [['=ARRAYFORMULA({"INS"; IF($K$6:K<>"", VLOOKUP($K$6:K,Lookups!$D$4:$E$27, 2, False), "")})']] }
  });

  const currentDateValues = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!B3`
  });
  const currentDate = currentDateValues.data.values[0][0];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!B6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[currentDate]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!C6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[userIdValue]] }
  });

  const userSheetName = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `Member Data!B${targetRowIndex1+1}:F${targetRowIndex1+1}`
  });
  const userSheetNameValues = userSheetName.data.values;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!D6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[userSheetNameValues[0][1]]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!E6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[userSheetNameValues[0][2]]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!G6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[userSheetNameValues[0][4]]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!H5`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [['=ARRAYFORMULA({"INS"; IF($G$6:G<>"", VLOOKUP($G$6:G,Lookups!$D$4:$E$27, 2, False), "")})']] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!I6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[action]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!K6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[rank]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!J5`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [['=ARRAYFORMULA({"INS"; IF($K$6:K<>"", VLOOKUP($K$6:K,Lookups!$D$4:$E$27, 2, False), "")})']] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!M6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[note]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!O6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[coUserIdValue]] }
  });

  const coUserSheetName = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `Member Data!C${targetRowIndex2+1}`
  });
  const coUserSheetNameValue = coUserSheetName.data.values[0][0];
  
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Rank Change!P6`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[coUserSheetNameValue]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Member Data!F${targetRowIndex1+1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[rank]] }
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Member Data!L${targetRowIndex1+1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[currentDate]] }
  });
  
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Member Data!M${targetRowIndex2+1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { majorDimension: "ROWS", values: [[currentDate]] }
  });

  /*
  if (discordIndex === -1 || rankIndex === -1) {
    return "❌ Could not find 'Discord ID' or 'Rank' columns.";
  }

  // Find row
  const targetRowIndex = rows.findIndex(r => r[idIndex] == robloxId);
  if (targetRowIndex === -1) {
    return `❌ User ${username} not found.`;
  }
  */
  let newRank = rank;
  /*
  // Auto logic if no rank is provided
  if (!newRank) {
    const currentRank = rows[targetRowIndex][rankIndex] || "";

    if (action === "promote") {
      newRank = currentRank + " +"; // Placeholder: replace with lookup logic
    } else if (action === "demote") {
      newRank = currentRank + " -"; // Placeholder
    } else if (action === "transfer") {
      newRank = currentRank + " (Transferred)"; // Placeholder
    }
  }

  // Update Google Sheet
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `Member Data!${String.fromCharCode(65 + rankIndex)}${targetRowIndex + 1}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[newRank]] }
  });*/
  return `✅ ${action} successful. <@${user.id}> is now ${newRank}.`;
}
