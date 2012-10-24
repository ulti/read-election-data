#include <string>
#include <list>
#include <vector>
#include <iostream>
#include <sstream>
#include <utility>

#include "tinyxml2.h"

class CDocumentProperties
{
public:
  std::string Title;
  std::string Author;
  std::string Created;
};

class CRegionProfile
{
public:
  std::string RegionName;
  int RegisteredVoters;
  int BallotsCast;
  double VoterTurnout;
};

class CLabeledTuple
{
public:
  std::string Label;
  std::vector<int> Data;
};

class CElectionHeader
{
public:
  std::string ColumnName;
  std::string CandidateName;
};

class CElection
{
public:
  std::string ElectionName;
  std::vector<CElectionHeader> Header;
  std::list<CLabeledTuple> Results;
};

class CScytlReader
{
public:
  CScytlReader(const std::string &Filename);
  ~CScytlReader();

  int Read();

protected:
  typedef std::pair<int,std::string> TTocEntry;

  int readDocumentProperties(const tinyxml2::XMLElement *dp, CDocumentProperties &documentProperties);
  int readTableOfContentsWorksheet(const tinyxml2::XMLElement *ws, std::list<TTocEntry> &toc);
  int readRegisteredVotersWorksheet(const tinyxml2::XMLElement *ws, std::list<CRegionProfile> &regionProfiles);
  int readElectionResultsWorksheet(const tinyxml2::XMLElement *ws, CElection &election);

private:
  std::string filename;
  tinyxml2::XMLDocument doc;

  CDocumentProperties documentProperties;
  std::list<TTocEntry> tableOfContents;
  std::list<CRegionProfile> regionProfiles;
  std::list<CElection> electionResults;
};

using namespace std;
using namespace tinyxml2;

CScytlReader::CScytlReader(const string &Filename)
  : filename(Filename)
{
}

CScytlReader::~CScytlReader()
{
}

int CScytlReader::readDocumentProperties(const XMLElement *dp, CDocumentProperties &documentProperties)
{
  if (!dp)
    return 1;

  const XMLElement *title = dp->FirstChildElement("o:Title");
  if (!title) return 1;
  documentProperties.Title = title->GetText();

  const XMLElement *author = dp->FirstChildElement("o:Author");
  if (!author) return 1;
  documentProperties.Author = author->GetText();

  const XMLElement *created = dp->FirstChildElement("o:Created");
  if (!created) return 1;
  documentProperties.Created = created->GetText();

  return 0;
}

int CScytlReader::readTableOfContentsWorksheet(const XMLElement *ws, list<TTocEntry> &toc)
{
  const XMLElement *table = ws->FirstChildElement("s:Table");
  if (!table)
    return 1;

  const XMLElement *row;
  for (row = table->FirstChildElement("s:Row"); row; row = row->NextSiblingElement())
  {
    // the Table of Contents entries we're interested in will have two cells on the same row,
    // the first will be a Number and the second will be a String. anything else is not
    // important to us and can be ignored.

    // Example:
    //
    //  <s:Row>
    //    <s:Cell s:StyleID="Page">
    //      <s:Data s:Type="Number">1</s:Data>
    //    </s:Cell>
    //    <s:Cell>
    //      <s:Data s:Type="String">Registered Voters</s:Data>
    //    </s:Cell>
    //  </s:Row>

    const XMLElement *cell1 = row->FirstChildElement("s:Cell");
    const XMLElement *cell2 = cell1 ? cell1->NextSiblingElement() : NULL;
    if (!cell1 || !cell2)
      continue;
    {
      const char *cellstyle = cell1->Attribute("s:StyleID");
      if (!cellstyle || strcmp(cellstyle, "Page"))
        continue;
    }

    const XMLElement *data1 = cell1->FirstChildElement("s:Data");
    const XMLElement *data2 = cell2->FirstChildElement("s:Data");
    if (!data1 || !data2)
      continue;

    if (strcmp(data1->Attribute("s:Type"), "Number") || 
        strcmp(data2->Attribute("s:Type"), "String"))
      continue;

    int page;
    if (data1->QueryIntText(&page) != XML_SUCCESS)
      continue;

    string contest = data2->GetText();

    toc.push_back(TTocEntry(page, contest));
  }

  return 0;
}

int CScytlReader::readRegisteredVotersWorksheet(const XMLElement *ws, list<CRegionProfile> &regionProfiles)
{
  const XMLElement *table = ws->FirstChildElement("s:Table");
  if (!table)
    return 1;

  // the first row contains our header

  // Example:
  //
  //  <s:Row>
  //    <s:Cell>
  //      <s:Data s:Type="String">County</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Registered Voters</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Ballots Cast</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Voter Turnout</s:Data>
  //    </s:Cell>
  //  </s:Row>

  list<string> header;
  const XMLElement *row = table->FirstChildElement("s:Row");
  if (row)
  {
    for (const XMLElement *cell = row->FirstChildElement("s:Cell");
         cell;
         cell = cell->NextSiblingElement())
    {
      const XMLElement *data = cell->FirstChildElement("s:Data");
      if (data)
      {
        const char *type = data->Attribute("s:Type");
        if (!strcmp(type, "String"))
          header.push_back(data->GetText());
      }
    }
    row = row->NextSiblingElement();
  }

  // now, read in voter data one row at a time. because this is the
  // registered voters page, we know what columns we should expect.

  // Example:
  //
  //  <s:Row>
  //    <s:Cell>
  //      <s:Data s:Type="String">Arkansas</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">9095</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">1898</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="String">20.87 %</s:Data>
  //    </s:Cell>
  //  </s:Row>

  for (;
       row;
       row = row->NextSiblingElement())
  {
    CRegionProfile profile;

    const XMLElement *cell = row->FirstChildElement("s:Cell");

    // read the region name (aka county/precinct name)
    if (cell)
    {
      const XMLElement *data = cell->FirstChildElement("s:Data");
      const char *type = data ? data->Attribute("s:Type") : NULL;
      if (type && !strcmp(type, "String"))
        profile.RegionName = data->GetText();
      cell = cell->NextSiblingElement();
    }

    // Use header to determine remaining columns
    list<string>::iterator itHeader = header.begin();
    // ignore first header entry (County) because it may differ for precinct-level files
    if (itHeader != header.end()) ++itHeader;

    for (;
         cell && itHeader != header.end();
         cell = cell->NextSiblingElement(), ++itHeader)
    {
      const char *style = cell->Attribute("s:StyleID");
      const XMLElement *data = cell->FirstChildElement("s:Data");
      const char *type = data ? data->Attribute("s:Type") : NULL;

      // we'll use the header name, the cell style, and the data type to verify file integrity
      // and make sure we're reading from the correct column.
      if (*itHeader == "Registered Voters"
          && style && !strcmp(style, "VoteCount")
          && type && !strcmp(type, "Number"))
      {
        if (data->QueryIntText(&profile.RegisteredVoters) != XML_SUCCESS)
          return 1;
      }
      else if (*itHeader == "Ballots Cast"
               && style && !strcmp(style, "VoteCount")
               && type && !strcmp(type, "Number"))
      {
        if (data->QueryIntText(&profile.BallotsCast) != XML_SUCCESS)
          return 1;
      }
      else if (*itHeader == "Voter Turnout"
               && style && !strcmp(style, "VoteCount")
               && type && !strcmp(type, "String"))
      {
        string turnout_str = data->GetText();
        // truncate to remove the appended percent sign
        istringstream buf(turnout_str.substr(0, turnout_str.length()-2));

        // convert to double, and make sure we read the WHOLE string
        buf >> profile.VoterTurnout;
        if (buf.fail()) return 1;
        buf.peek();
        if (!buf.eof()) return 1;
      }
      else
      {
        cout << "Error: unrecognized column name in Registered Voters worksheet (header says '" << *itHeader << "')" << endl;
        return 1;
      }
    }

    // make sure we have the same number of headers and columns
    if (cell || itHeader != header.end())
      return 1;

    regionProfiles.push_back(profile);
  }

  return 0;
}

int CScytlReader::readElectionResultsWorksheet(const XMLElement *ws, CElection &election)
{
  const XMLElement *table = ws->FirstChildElement("s:Table");
  if (!table)
    return 1;

  // first row should contain our election name

  // Example:
  //
  //  <s:Row>
  //    <s:Cell s:MergeAcross="6" s:StyleID="headerLbl">
  //      <s:Data s:Type="String">U.S. President - DEM</s:Data>
  //    </s:Cell>
  //  </s:Row>

  const XMLElement *row = table->FirstChildElement("s:Row");
  {
    const XMLElement *cell = row ? row->FirstChildElement("s:Cell") : NULL;
    const XMLElement *data = cell ? cell->FirstChildElement("s:Data") : NULL;
    const char *style = cell ? cell->Attribute("s:StyleID") : NULL;
    const char *type = data ? data->Attribute("s:Type") : NULL;
    if (style && !strcmp(style, "headerLbl") &&
        type && !strcmp(type, "String"))
    {
      election.ElectionName = data->GetText();

      // use MergeAcross to determine how many columns there are
      // so we can allocate the header's vector
      int mergeacross = 0;
      cell->QueryIntAttribute("s:MergeAcross", &mergeacross);
      election.Header.resize(mergeacross + 1);

      row = row->NextSiblingElement();
    }
    else return 1;
  }

  // next row has candidate names. be careful about the "MergeAcross" attribute.

  // Example:
  //
  //  <s:Row>
  //    <s:Cell>
  //      <s:Data s:Type="String"/>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String"/>
  //    </s:Cell>
  //    <s:Cell s:MergeAcross="1">
  //      <s:Data s:Type="String">John Wolfe</s:Data>
  //    </s:Cell>
  //    <s:Cell s:MergeAcross="1">
  //      <s:Data s:Type="String">Barack Obama</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String"/>
  //    </s:Cell>
  //  </s:Row>

  {
    vector<CElectionHeader>::iterator itHeader = election.Header.begin();
    const XMLElement *cell = row ? row->FirstChildElement("s:Cell") : NULL;
    for (;
         cell && itHeader != election.Header.end();
         cell = cell->NextSiblingElement())
    {
      const XMLElement *data = cell ? cell->FirstChildElement("s:Data") : NULL;
      int mergeacross = 0;
      cell->QueryIntAttribute("s:MergeAcross", &mergeacross);

      for (int i = 0;
           i <= mergeacross && itHeader != election.Header.end();
           ++i, ++itHeader)
      {
        if (data && data->GetText())
          itHeader->CandidateName = data->GetText();
      }
    }

    // make sure we used the right number of columns
    if (cell || itHeader != election.Header.end())
      return 1;

    row = row->NextSiblingElement();
  }

  // the next row has column names

  // Example:
  //
  //  <s:Row>
  //    <s:Cell>
  //      <s:Data s:Type="String">County</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Registered Voters</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Election Day</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Total Votes</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Election Day</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Total Votes</s:Data>
  //    </s:Cell>
  //    <s:Cell>
  //      <s:Data s:Type="String">Total</s:Data>
  //    </s:Cell>
  //  </s:Row>

  {
    vector<CElectionHeader>::iterator itHeader = election.Header.begin();
    const XMLElement *cell = row ? row->FirstChildElement("s:Cell") : NULL;
    for (;
         cell && itHeader != election.Header.end();
         cell = cell->NextSiblingElement(), ++itHeader)
    {
      const XMLElement *data = cell->FirstChildElement("s:Data");
      const char *type = data ? data->Attribute("s:Type") : NULL;

      if (type && !strcmp(type, "String") && data->GetText())
        itHeader->ColumnName = data->GetText();
      else
        return 1;
    }

    // make sure we used the right number of columns
    if (cell || itHeader != election.Header.end())
      return 1;

    row = row->NextSiblingElement();
  }

  // the rest of the rows are voter data

  // Example:
  //
  //  <s:Row>
  //    <s:Cell>
  //      <s:Data s:Type="String">Arkansas</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">0</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">508</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">508</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">599</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">599</s:Data>
  //    </s:Cell>
  //    <s:Cell s:StyleID="VoteCount">
  //      <s:Data s:Type="Number">1107</s:Data>
  //    </s:Cell>
  //  </s:Row>

  for (;
       row;
       row = row->NextSiblingElement())
  {
    CLabeledTuple tuple;

    // the first cell is the label (a string)
    const XMLElement *cell = row->FirstChildElement("s:Cell");
    {
      const XMLElement *data = cell ? cell->FirstChildElement("s:Data") : NULL;
      const char *type = data ? data->Attribute("s:Type") : NULL;
      if (type && !strcmp(type, "String"))
        tuple.Label = data->GetText();
      else return 1;

      cell = cell->NextSiblingElement();
    }

    // the remaining cells are vote counts (integers)
    for (;
         cell;
         cell = cell->NextSiblingElement())
    {
      const XMLElement *data = cell ? cell->FirstChildElement("s:Data") : NULL;
      const char *style = cell ? cell->Attribute("s:StyleID") : NULL;
      const char *type = data ? data->Attribute("s:Type") : NULL;
      if (style && !strcmp(style, "VoteCount") && type && !strcmp(type, "Number"))
      {
        int value;
        if (data->QueryIntText(&value) != XML_SUCCESS)
          return 1;
        tuple.Data.push_back(value);
      } else return 1;
    }

    // make sure our tuple has the right length
    if (tuple.Data.size() + 1 != election.Header.size())
      return 1;

    election.Results.push_back(tuple);
  }

  return 0;
}

int CScytlReader::Read()
{
  doc.LoadFile(filename.c_str());
  if (doc.Error()) {
    doc.PrintError();
    cout << "Error loading <" << filename << ">" << endl;
    return 1;
  }

  // locate root node
  const XMLElement *root = doc.FirstChildElement("s:Workbook");
  if (!root) {
    cout << "Couldn't find root s:Workbook node" << endl;
    return 1;
  }

  // read document properties
  const XMLElement *dp = root->FirstChildElement("o:DocumentProperties");
  if (readDocumentProperties(dp, documentProperties)) {
    cout << "Error reading document properties" << endl;
    return 1;
  }

  // build table of contents
  const XMLElement *ws = root->FirstChildElement("s:Worksheet");
  while (ws && strcmp(ws->Attribute("s:Name"), "Table of Contents"))
    ws = ws->NextSiblingElement();

  if (readTableOfContentsWorksheet(ws, tableOfContents)) {
    cout << "Error reading table of contents" << endl;
    return 1;
  }

  // read registered voter info
  while (ws && strcmp(ws->Attribute("s:Name"), "Registered Voters"))
    ws = ws->NextSiblingElement();
  if (readRegisteredVotersWorksheet(ws, regionProfiles)) {
    cout << "Error reading registered voters worksheet" << endl;
    return 1;
  }

  // read election info
  ws = ws->NextSiblingElement();
  for (;
       ws;
       ws = ws->NextSiblingElement())
  {
    CElection election;
    if (readElectionResultsWorksheet(ws, election)) {
      cout << "Error reading election results worksheet" << endl;
      return 1;
    }
    electionResults.push_back(election);
  }

  // test document properties
  cout << "Title;" << documentProperties.Title << endl
       << "Author;" << documentProperties.Author << endl
       << "Created;" << documentProperties.Created << endl;

  // test table of contents
  for (list<TTocEntry>::const_iterator tocIt = tableOfContents.begin();
       tocIt != tableOfContents.end();
       ++tocIt)
  {
    cout << tocIt->first << ";" << tocIt->second << endl;
  }

  // test registered voters
  cout << "County;Registered Voters;Ballots Cast;Voter Turnout" << endl;
  for (list<CRegionProfile>::const_iterator itRegion = regionProfiles.begin();
       itRegion != regionProfiles.end();
       ++itRegion)
  {
    cout << "  " << itRegion->RegionName << ";"
                 << itRegion->RegisteredVoters << ";"
                 << itRegion->BallotsCast << ";"
                 << itRegion->VoterTurnout << endl;
  }

  // test election results
  for (list<CElection>::const_iterator itElection = electionResults.begin();
       itElection != electionResults.end();
       ++itElection)
  {
    cout << itElection->ElectionName << endl;

    for (vector<CElectionHeader>::const_iterator itHeader = itElection->Header.begin();
         itHeader != itElection->Header.end();
         ++itHeader)
    {
      if (itHeader != itElection->Header.begin())
        cout << ";";

      if (itHeader->CandidateName != "")
        cout << itHeader->CandidateName << " - ";
      cout << itHeader->ColumnName;
    }
    cout << endl;

    for (list<CLabeledTuple>::const_iterator itTuple = itElection->Results.begin();
         itTuple != itElection->Results.end();
         ++itTuple)
    {
      cout << itTuple->Label << ";";
      for (vector<int>::const_iterator itData = itTuple->Data.begin();
           itData != itTuple->Data.end();
           ++itData)
      {
        if (itData != itTuple->Data.begin())
          cout << ";";
        cout << *itData;
      }
      cout << endl;
    }
  }

  return 0;
}

void usage(int argc, char * const *argv)
{
  cout << argv[0] << " <filename>" << endl;
}

int main(int argc, char **argv)
{
  string infile;

  int narg = 1;
  while (narg < argc)
  {
    // insert other arguments here

    infile = argv[narg++];
    break;
  }

  if (narg != argc || infile == "")
  {
    usage(argc, argv);
    exit(1);
  }

  CScytlReader fin(infile);
  if (fin.Read())
  {
    cout << "Error reading from <" << infile << ">" << endl;
    return 1;
  }

  return 0;
}
